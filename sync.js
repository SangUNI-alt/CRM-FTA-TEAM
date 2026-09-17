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

const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

function getSharingUrlToken(url) {
  const base64Value = Buffer.from(url).toString('base64');
  return "u!" + base64Value.replace(/=/g, '').replace(/\//g, '_').replace(/\+/g, '-');
}

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
  if (data.error) throw new Error("Lỗi lấy Token từ Microsoft: " + data.error_description);
  return data.access_token;
}

async function runSync() {
  console.log("🚀 Bắt đầu quá trình quét và đồng bộ dữ liệu...");
  
  try {
    const token = await getAccessToken();
    const shareToken = getSharingUrlToken(PARENT_FOLDER_ID);
    
    let url = `https://graph.microsoft.com/v1.0/shares/${shareToken}/driveItem/children?$top=200`;
    let itemsToInsert = [];

    while (url) {
      console.log(`Đang tải dữ liệu từ Graph API... (Đã quét được ${itemsToInsert.length} items)`);
      const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
      const data = await res.json();

      if (data.error) throw new Error(data.error.message);

      if (data.value) {
        data.value.forEach(item => {
          itemsToInsert.push({
            id: item.id,
            name: item.name,
            is_folder: !!item.folder,
            parent_id: item.parentReference ? item.parentReference.id : null,
            web_url: item.webUrl,
            size: item.size || 0,
            created_at: item.createdDateTime
          });
        });
      }
      url = data['@odata.nextLink'] || null;
    }

    console.log(` Quét xong ${itemsToInsert.length} items. Đang nạp vào Supabase...`);

    const { error } = await supabase.from('ecis_items').upsert(itemsToInsert);

    if (error) {
      console.error("❌ Lỗi Supabase:", error.message);
    } else {
      console.log(`✅ ĐỒNG BỘ THÀNH CÔNG! Đã cập nhật ${itemsToInsert.length} items vào Database.`);
    }

  } catch (err) {
    console.error("❌ Lỗi chương trình:", err.message);
    process.exit(1);
  }
}

runSync();
