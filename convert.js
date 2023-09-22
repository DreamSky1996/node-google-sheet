const fs = require('fs');
const { utils, read } = require('xlsx');

const parseBaRawExcel = () => {
  const workbook = read('excel_data/ba.xlsx', { type: 'file' });
  const sheet_name_list = workbook.SheetNames[0];
  // console.log(workbook.Sheets[sheet_name_list]);
  const rows = utils.sheet_to_json(workbook.Sheets[sheet_name_list], {
    raw: true,
    dateNF: 'MM-DD-YYYY',
    header: 1,
    defval: '',
    range: 5,
  });
  // console.log(rows);
  const body = rows.slice(1, rows.length);
  // fs.writeFileSync('parse_excel_data/ba-result.json', JSON.stringify(rows));
  fs.writeFileSync('parse_excel_data/ba-result.json', JSON.stringify(body));
};

const parseAttributesRawExcel = (base64Excel) => {
  const workbook = read('excel_data/at.xlsx', { type: 'file' });
  const sheet_name_list = workbook.SheetNames[0];
  // console.log(workbook.Sheets[sheet_name_list]);
  const rows = utils.sheet_to_json(workbook.Sheets[sheet_name_list], {
    raw: true,
    dateNF: 'MM-DD-YYYY',
    header: 1,
    defval: '',
  });
  // console.log(rows);
  const headers = rows[0];

  const body = rows.slice(4, rows.length);
  // fs.writeFileSync('parse_excel_data/attributes-result.json', JSON.stringify([headers, ...body]));
  fs.writeFileSync('parse_excel_data/attributes-result.json', JSON.stringify(body));
};

const parseAdPromotedProductRawExcel = (base64Excel) => {
  const workbook = read('excel_data/ad.xls', { type: 'file' });
  const sheet_name_list = workbook.SheetNames[0];
  // console.log(workbook.Sheets[sheet_name_list]);
  const rows = utils.sheet_to_json(workbook.Sheets[sheet_name_list], {
    raw: true,
    dateNF: 'MM-DD-YYYY',
    header: 1,
    defval: '',
    range: 7,
  });
  fs.writeFileSync('parse_excel_data/ad-result.json', JSON.stringify(rows));
};

parseBaRawExcel();
parseAttributesRawExcel();
// parseAdPromotedProductRawExcel();