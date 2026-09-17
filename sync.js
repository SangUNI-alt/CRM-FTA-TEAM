const { createClient } = require('@supabase/supabase-js');

// =========================================================
// THÔNG TIN CẤU HÌNH
// =========================================================
const TENANT_ID = "a70e2894-d043-4c2d-b843-b376e7e7df4b";
const CLIENT_ID = "1f5c6852-079c-463a-8fc8-e814359698ee";
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const PARENT_FOLDER_ID = "https://uniconsulting079.sharepoint.com/:f:/s/SEDOVINA_CO/IgAJaovqxuNDQq6KUumP_IkHAXQkx84t9aAqB-3t5UwndeE?e=OEWhSl";

const SUPABASE_URL = "https://eggshqsdtieqzjtttdyh.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImVnZ3NocXNkcmllcXpqdHR0ZHloIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODk2MjQ5MDMsImV4cCI6MjEwNTIwMDkwM30.7hgv-vm2GDL3xEa3IvFWfrtynCzWWlWDBDgLMrbuKw4";

const supabase = createClient(SUPABASE_URL, SUPABASE_KEY, {
  auth: { persistSession: false }
});

function getSharingUrlToken(url) {
  const base64Value = Buffer.from(url).toString('base64');
  return "u!" + base64Value.replace(/=/g, '').replace(/\//g, '_').replace(/\+/g, '-');
}

// 1. Lấy Access Token từ Microsoft Graph
async function getAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });

  const res = await fetch(tokenUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params
  });

  const data = await res.json();
  if (data.error) throw new Error("Lỗi lấy Token: " + data.error_description);
  return data.access_token;
}

// 2. Hàm hỗ trợ lấy danh sách con của 1 folder từ Graph API
async function getChildren(token, shareToken, driveId, itemId) {
  let url = (driveId && itemId)
    ? `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/children?$top=200`
    : `https://graph.microsoft.com/v1.0/shares/${shareToken}/driveItem/children?$top=200`;

  let items = [];
  while (url) {
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    const data = await res.json();
    if (data.error) throw new Error(`Lỗi lấy dữ liệu folder: ${data.error.message}`);
    if (data.value) items.push(...data.value);
    url = data['@odata.nextLink'] || null;
  }
  return items;
}

// 3. Tiến hành quét phân cấp theo đúng Logic yêu cầu
async function runSync() {
  console.log("🚀 Bắt đầu quá trình lọc và đồng bộ dữ liệu...");
  
  try {
    const token = await getAccessToken();
    const shareToken = getSharingUrlToken(PARENT_FOLDER_ID);
    
    // Đọc các mục tại gốc (thư mục 2026)
    const rootItems = await getChildren(token, shareToken);
    
    // Tìm 2 folder con FCL và LCL
    let level1Folders = rootItems.filter(item => item.folder && ["FCL", "LCL"].includes(item.name.toUpperCase()));

    // Trường hợp link chia sẻ chứa folder 2026 ở bên trong
    if (level1Folders.length === 0) {
      const folder2026 = rootItems.find(item => item.folder && item.name === "2026");
      if (folder2026) {
        const items2026 = await getChildren(token, shareToken, folder2026.parentReference.driveId, folder2026.id);
        level1Folders = items2026.filter(item => item.folder && ["FCL", "LCL"].includes(item.name.toUpperCase()));
      }
    }

    if (level1Folders.length === 0) {
      console.log("⚠️ Không tìm thấy folder FCL hoặc LCL nào.");
      return;
    }

    let recordsToUpsert = [];

    // Duyệt qua từng folder FCL và LCL
    for (const l1Folder of level1Folders) {
      const categoryName = l1Folder.name.toUpperCase();
      console.log(`\n📂 Đang quét mục: ${categoryName}`);

      const driveId = l1Folder.parentReference.driveId;
      const l2Items = await getChildren(token, shareToken, driveId, l1Folder.id);
      
      // LẤY ĐÚNG 5 FOLDER CON 2 ĐẦU TIÊN
      const l2Folders = l2Items.filter(item => item.folder).slice(0, 5);
      console.log(` Found ${l2Folders.length} folder con 2 đầu tiên.`);

      for (const l2Folder of l2Folders) {
        console.log(`  └─ Quét Folder Con 2: "${l2Folder.name}"`);
        
        // Đọc các mục bên trong Folder Con 2
        const l3Items = await getChildren(token, shareToken, driveId, l2Folder.id);

        // 1. Lọc tất cả file PDF có trong Folder Con 2
        const pdfFiles = l3Items
          .filter(item => !item.folder && item.name.toLowerCase().endsWith('.pdf'))
          .map(item => item.name);

        // 2. Lấy các Folder Con 3 để quét file XLS/XLSX
        const l3Folders = l3Items.filter(item => item.folder);
        let xlsFiles = [];

        for (const l3Folder of l3Folders) {
          const l4Items = await getChildren(token, shareToken, driveId, l3Folder.id);
          const xlInL3 = l4Items
            .filter(item => !item.folder && (item.name.toLowerCase().endsWith('.xls') || item.name.toLowerCase().endsWith('.xlsx')))
            .map(item => item.name);
          
          xlsFiles.push(...xlInL3);
        }

        // Đóng gói bản ghi
        recordsToUpsert.push({
          id: l2Folder.id,
          category: categoryName,
          folder_name: l2Folder.name,
          pdf_count: pdfFiles.length,
          pdf_list: pdfFiles,
          xls_count: xlsFiles.length,
          xls_list: xlsFiles,
          updated_at: new Date().toISOString()
        });
      }
    }

    console.log(`\n Nạp ${recordsToUpsert.length} bản ghi vào bảng 'ecis_summary' trên Supabase...`);

    // Ghi dữ liệu vào bảng ecis_summary
    const { error } = await supabase.from('ecis_summary').upsert(recordsToUpsert);

    if (error) {
      console.error("❌ Lỗi Supabase:", error.message);
    } else {
      console.log("✅ ĐỒNG BỘ THÀNH CÔNG DỮ LIỆU TỔNG HỢP VÀO SUPABASE!");
    }

  } catch (err) {
    console.error("❌ Lỗi tiến trình:", err.message);
    process.exit(1);
  }
}

runSync();
