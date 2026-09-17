const { createClient } = require('@supabase/supabase-js');

// =========================================================
// THÔNG TIN CẤU HÌNH
// =========================================================
const TENANT_ID = "a70e2894-d043-4c2d-b843-b376e7e7df4b";
const CLIENT_ID = "1f5c6852-079c-463a-8fc8-e814359698ee";
const CLIENT_SECRET = process.env.CLIENT_SECRET;

const SITE_HOST = "uniconsulting079.sharepoint.com";
const SITE_PATH = "/sites/SEDOVINA_CO";

const SUPABASE_URL = "https://eggshqsdtieqzjtttdyh.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImVnZ3NocXNkcmllcXpqdHR0ZHloIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODk2MjQ5MDMsImV4cCI6MjEwNTIwMDkwM30.7hgv-vm2GDL3xEa3IvFWfrtynCzWWlWDBDgLMrbuKw4";

const supabase = createClient(SUPABASE_URL, SUPABASE_KEY, {
  auth: { persistSession: false }
});

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

// 2. Lấy danh sách items con trong 1 folder
async function getChildren(token, driveId, itemId = 'root') {
  let url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/children?$top=200`;
  let items = [];

  while (url) {
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    const data = await res.json();
    if (data.error) throw new Error(`Lỗi API (${itemId}): ${data.error.message}`);
    if (data.value) items.push(...data.value);
    url = data['@odata.nextLink'] || null;
  }
  return items;
}

// 3. Tiến hành quét phân cấp
async function runSync() {
  console.log("🚀 Bắt đầu quá trình kết nối SharePoint và lọc dữ liệu...");
  
  try {
    const token = await getAccessToken();

    // Lấy Site ID & Default Drive ID
    const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${SITE_HOST}:${SITE_PATH}`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    const siteData = await siteRes.json();
    if (siteData.error) throw new Error("Lỗi lấy Site ID: " + siteData.error.message);

    const driveRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    const driveData = await driveRes.json();
    if (driveData.error) throw new Error("Lỗi lấy Drive ID: " + driveData.error.message);

    const driveId = driveData.id;
    console.log(`✅ Kết nối thành công Site SharePoint! Drive ID: ${driveId}`);

    // Đọc gốc Document Library
    let rootItems = await getChildren(token, driveId, 'root');

    // Tìm thư mục "2026" nếu nằm ở cấp gốc
    const folder2026 = rootItems.find(item => item.folder && item.name === "2026");
    if (folder2026) {
      console.log("📁 Đã tìm thấy thư mục '2026', tiến hành đi vào bên trong...");
      rootItems = await getChildren(token, driveId, folder2026.id);
    }

    // Lọc ra FCL và LCL
    const level1Folders = rootItems.filter(item => item.folder && ["FCL", "LCL"].includes(item.name.toUpperCase()));

    if (level1Folders.length === 0) {
      console.log("⚠️ Không tìm thấy folder FCL hoặc LCL nào trong danh sách.");
      return;
    }

    let recordsToUpsert = [];

    // Duyệt FCL và LCL
    for (const l1Folder of level1Folders) {
      const categoryName = l1Folder.name.toUpperCase();
      console.log(`\n📂 Đang quét mục: ${categoryName}`);

      const l2Items = await getChildren(token, driveId, l1Folder.id);
      
      // LẤY ĐÚNG 5 FOLDER CON 2 ĐẦU TIÊN
      const l2Folders = l2Items.filter(item => item.folder).slice(0, 5);
      console.log(` Found ${l2Folders.length} folder con 2 đầu tiên.`);

      for (const l2Folder of l2Folders) {
        console.log(`  └─ Quét Folder Con 2: "${l2Folder.name}"`);
        
        const l3Items = await getChildren(token, driveId, l2Folder.id);

        // Lọc file PDF ở Folder Con 2
        const pdfFiles = l3Items
          .filter(item => !item.folder && item.name.toLowerCase().endsWith('.pdf'))
          .map(item => item.name);

        // Quét các Folder Con 3 để tìm file XLS/XLSX
        const l3Folders = l3Items.filter(item => item.folder);
        let xlsFiles = [];

        for (const l3Folder of l3Folders) {
          const l4Items = await getChildren(token, driveId, l3Folder.id);
          const xlInL3 = l4Items
            .filter(item => !item.folder && (item.name.toLowerCase().endsWith('.xls') || item.name.toLowerCase().endsWith('.xlsx')))
            .map(item => item.name);
          
          xlsFiles.push(...xlInL3);
        }

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
