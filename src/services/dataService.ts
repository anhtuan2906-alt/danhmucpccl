import Papa from 'papaparse';

export const SHEET_GIDS: Record<string, string> = {
  'QLĐT': '1281741019',
  'KTAT': '30725636',
  'KHVT': '1833969714',
  'TCG': '1934956086',
  'TVTK': '845682162',
  'TVGS': '167241815',
  'XL': '1625219412',
  'TCKT': '1019933485',
  'QLLĐ': '2016902786',
  'DVĐL': '838719338',
  'ĐĐHTĐ': '1220487931',
  'ETC': '1696506223',
  'Kiểm toán': '354111578',
  'SXD': '605843879',
  'TCT': '1094941156',
  'UBND': '716705477'
};

const SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/14BF0RUfBq-Arl6ngVvD44fQnNayBEC1Xtz-RFgzA4GI/export?format=csv&gid=0';
const PROXY_URL = `https://corsproxy.io/?${encodeURIComponent(SHEET_CSV_URL)}`;

export async function fetchWithFallback(url: string, isText = true) {
  try {
    const res = await fetch(url);
    if (res.ok) return isText ? await res.text() : res;
  } catch (e) {}

  try {
    const proxyUrl = `https://corsproxy.io/?${encodeURIComponent(url)}`;
    const res = await fetch(proxyUrl);
    if (res.ok) return isText ? await res.text() : res;
  } catch (e) {}
  
  try {
    const allOriginsUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`;
    const res = await fetch(allOriginsUrl);
    if (res.ok) return isText ? await res.text() : res;
  } catch (e) {}

  throw new Error("Cannot fetch: " + url);
}

export async function loadProjectData(projectCode: string, templateDataCache?: string[][]) {
  // 1. Fetch template data if not provided
  let templateRows = templateDataCache;
  if (!templateRows) {
    const csvText = await fetchWithFallback(`${SHEET_CSV_URL}&t=${Date.now()}`);
    const results = Papa.parse(csvText as string);
    const allRows = results.data as string[][];
    // Template header is at row 15 (index 14). Data starts at index 15.
    templateRows = allRows.slice(15).filter(row => row.some(cell => cell.trim() !== ''));
  }

  // 2. Identify unique departments
  const departments = new Set<string>();
  templateRows.forEach(row => {
    // row[6] is 'Cơ quan ban hành'
    const dept = row[6]?.trim();
    if (dept && SHEET_GIDS[dept]) {
      departments.add(dept);
    }
  });

  // 3. Fetch HTML view for each department to get data and hyperlinks
  const departmentData = new Map<string, any[]>();
  
  if (projectCode) {
    const fetchPromises = Array.from(departments).map(async (dept) => {
      const gid = SHEET_GIDS[dept];
      const htmlUrl = `https://docs.google.com/spreadsheets/d/14BF0RUfBq-Arl6ngVvD44fQnNayBEC1Xtz-RFgzA4GI/htmlview/sheet?headers=true&gid=${gid}&t=${Date.now()}`;
      try {
        const htmlText = await fetchWithFallback(htmlUrl) as string;
        const parser = new DOMParser();
        const doc = parser.parseFromString(htmlText, 'text/html');
        
        const rows = doc.querySelectorAll('table tbody tr');
        const docs: any[] = [];
        
        rows.forEach(row => {
          const cells = row.querySelectorAll('td');
          if (cells.length >= 6) {
            // Headers: STT, Mã CT, Nội dung văn bản, Số VB, Ngày VB, File văn bản, Ghi chú
            const maCT = cells[1].textContent?.trim();
            if (maCT === projectCode) {
              const noiDung = cells[2].textContent?.replace(/\s+/g, ' ').trim();
              const soVb = cells[3].textContent?.trim();
              const ngayVb = cells[4].textContent?.trim();
              
              // Extract hyperlink
              const fileCell = cells[5];
              let fileUrl = fileCell.textContent?.trim();
              const linkElement = fileCell.querySelector('a');
              if (linkElement && linkElement.href) {
                let href = linkElement.href;
                if (href.includes('google.com/url?')) {
                  try {
                    const urlParams = new URLSearchParams(href.split('?')[1]);
                    const q = urlParams.get('q');
                    if (q) href = q;
                  } catch (e) {}
                }
                fileUrl = href;
              }

              docs.push({
                noiDung,
                soVb,
                ngayVb,
                fileUrl
              });
            }
          }
        });
        
        departmentData.set(dept, docs);
      } catch (e) {
        console.warn(`Failed to fetch sheet ${dept}`, e);
      }
    });

    await Promise.allSettled(fetchPromises);
  }

  // 4. Map the fetched data back to the template rows
  const finalData = templateRows.map(row => {
    const newRow = [...row];
    const tenVanBan = row[1]?.replace(/\s+/g, ' ').trim();
    const dept = row[6]?.trim();
    
    if (dept && tenVanBan && departmentData.has(dept)) {
      const deptDocs = departmentData.get(dept);
      const matchedDoc = deptDocs?.find(doc => doc.noiDung === tenVanBan);
      
      if (matchedDoc) {
        newRow[3] = matchedDoc.soVb || '';
        newRow[4] = matchedDoc.ngayVb || '';
        newRow[5] = matchedDoc.fileUrl || '';
      } else {
        // Clear them if not found in the department sheet for this project
        newRow[3] = '';
        newRow[4] = '';
        newRow[5] = '';
      }
    } else if (tenVanBan) {
      // If no department or not found, clear the document specifics
      newRow[3] = '';
      newRow[4] = '';
      newRow[5] = '';
    }
    return newRow;
  });

  return { finalData, templateRows };
}
