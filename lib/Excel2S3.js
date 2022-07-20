const Excel2S3 = (function () {
  let AWS = require('aws-sdk');
  let Stream = require('stream');
  const exceljs = require('exceljs');
  let s3 = new AWS.S3();

  const setConfig = (config) => {
    if (config) {
      AWS.config.update({ region: config.AWS_DEFAULT_REGION })
      s3 = new AWS.S3({
        accessKeyId: config.AWS_ACCESS_KEY_ID,
        secretAccessKey: config.AWS_SECRET_ACCESS_KEY
      });
    }
  }

  const getColumnCode = (cellIndex, rowIndex) => {
    if (!cellIndex) throw new Error('The "cellIndex" is undefined.');
    if (!rowIndex) throw new Error('The "rowIndex" is undefined.');
    var cellIndex = cellIndex,
      rowIndex = rowIndex,
      alphabet = [
        'A', 'B', 'C', 'D', 'E', 'F', 'G',
        'H', 'I', 'J', 'K', 'L', 'M', 'N',
        'O', 'P', 'Q', 'R', 'S', 'T', 'U',
        'V', 'W', 'X', 'Y', 'Z'
      ],
      cell = '';
    if (cellIndex > alphabet.length) {
      var stacks = getCodeSet(cellIndex, [], alphabet);
      for (let ctr = stacks.length - 1; ctr >= 0; ctr--) {
        cell += alphabet[stacks[ctr] - 1];
      }
      cell += rowIndex;
    }
    else {
      cell = alphabet[cellIndex - 1] + rowIndex;
    }
    return cell;
  }

  const getCodeSet = (cellIndex, stacks, alphabet) => {
    var rem = cellIndex % alphabet.length;

    cellIndex = Math.floor(cellIndex / alphabet.length);

    if (rem === 0) {

      rem = alphabet.length;

      cellIndex -= 1;
    }

    stacks.push(rem);

    if (cellIndex > alphabet.length) {

      return getCodeSet(cellIndex, stacks, alphabet);
    }
    else {

      stacks.push(cellIndex);

      return stacks;
    }
  }

  class Excel2S3 {
    constructor(config) {
      setConfig(config);
      this[getColumnCode] = getColumnCode;
      this[getCodeSet] = getCodeSet;
      this.sheets = [];
    }
    writeBasicExcel = async (args) => {
      const headers = args.headers,
        worksheetName = args.worksheetName,
        excelRecordList = args.excelRecordList,
        bucket = args.bucket,
        key = args.key;

      var stream = new Stream.PassThrough();

      var workbook = new exceljs.Workbook();

      var worksheet = workbook.addWorksheet(worksheetName); //creating worksheet

      var sheetHeader = [];

      //  WorkSheet Header
      for (var ctr = 0; ctr < headers.length; ctr++) {
        var header = { header: headers[ctr], key: 'header' + ctr, width: 20 }
        sheetHeader.push(header)
      }

      worksheet.columns = sheetHeader;

      var font = {
        bold: true,
        color: { argb: 'FFFFFFFF' }
      }
      var fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF538DD5' }
      }

      for (var ctr = 0; ctr < headers.length; ctr++) {

        worksheet.getCell(String.fromCharCode(65 + ctr) + '1').font = font

        worksheet.getCell(String.fromCharCode(65 + ctr) + '1').fill = fill
      }

      // Add Array Rows
      worksheet.addRows(excelRecordList);

      return await workbook.xlsx.write(stream)
        .then(() => {
          return s3.upload({
            Key: key,
            Bucket: bucket,
            Body: stream,
            ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          }).promise();
        })
    }

    writeComplexExcel = async (args) => {
      const filename = args.filename,
        bucket = args.bucket;

      var stream = new Stream.PassThrough();
      var workbook = new exceljs.Workbook();

      // Iterate all the sheets
      for (var ctr1 = 0; ctr1 < this.sheets.length; ctr1++) {
        const sheet = this.sheets[ctr1];
        // Create work sheet
        const worksheet = workbook.addWorksheet(sheet.title, sheet.option);

        var columns = [];
        var headerKeys = [];
        var mainHeader;
        // Get the total number of header to reference where should the row start
        const ref = sheet.headers.length;
        // print headers in excel
        for (var ctr2 = 0; ctr2 < sheet.headers.length; ctr2++) {
          // add 1 because excel index start with 1
          const rowIndex = ctr2 + 1;
          // Get all the headers
          const headerList = sheet.headers[ctr2].headerNames;
          // Separate main header when it comes in rendering
          if (sheet.headers[ctr2].main) {
            // get row for columns
            const row = worksheet.getRow(ref);
            // store main header for setting other configuration
            mainHeader = sheet.headers[ctr2];
            // This will be the cursor for pointing where the settings and value should be put.
            var colIndex = 1;
            // iterate the column headers
            for (var ctr3 = 0; ctr3 < headerList.length; ctr3++) {
              if (headerList[ctr3].colspan > 1) {
                // if we have a colspan, we have to calculate the column code (eg.A1) where the merge start.
                var currentCell = this[getColumnCode](colIndex, rowIndex);
                // and the calculate the end code where is the end column of merge.
                var cellSpan = this[getColumnCode]((headerList[ctr3].colspan - 1) + colIndex, rowIndex);
                // and then merge
                worksheet.mergeCells(currentCell + ":" + cellSpan);
              }
              columns.push(headerList[ctr3].title);
              // add column settings
              headerKeys.push({ key: headerList[ctr3].key, width: headerList[ctr3].width });

              colIndex++;
              // because we merge cell we need to move the cursor to the 
              colIndex += headerList[ctr3].colspan - 1;
            }
            row.values = columns;
            worksheet.columns = headerKeys;
            // set header height
            row.height = mainHeader.height;
          }
          else {
            const row = worksheet.getRow(rowIndex);
            row.height = sheet.headers[ctr2].height
            let colIndex = 1;
            // ITERATE NAMES AND SET VALUES IN EACH CELL
            for (var ctr4 = 0; ctr4 < headerList.length; ctr4++) {
              const cell = row.getCell(colIndex);
              if (headerList[ctr4].colspan > 1) {
                let currentCell = this[getColumnCode](colIndex, rowIndex);
                let cellSpan = this[getColumnCode]((headerList[ctr4].colspan - 1) + colIndex, rowIndex);
                worksheet.mergeCells(currentCell + ":" + cellSpan);
              }
              cell.value = headerList[ctr4].title;
              const style = headerList[ctr4].style;
              for (const [key, value] of Object.entries(style)) {
                cell[key] = value;
              }

              colIndex++;
              colIndex += headerList[ctr4].colspan - 1;
            }
          }
        }
        // Add Row data
        worksheet.addRows(sheet.rowData.rows);

        worksheet.headerFooter = sheet.headerFooter;

        worksheet.eachRow(function (row, rowNumber) {

          if (rowNumber >= ref) {

            row.eachCell({ includeEmpty: true }, function (cell, cellNumber) {

              if (ref === rowNumber) {
                // add design to the header
                const style = mainHeader.headerNames[cellNumber - 1].style;

                for (const [key, value] of Object.entries(style)) {

                  cell[key] = value;
                }
              }
              else {

                const style = sheet.rowData.style;

                for (const [key, value] of Object.entries(style)) {

                  cell[key] = value;
                }
              }
            });
          }
        });

        // Display Footer

        if (sheet.footers.length > 0) {

          const footerRowStartIndex = ref + sheet.rowData.rows.length;

          const maxFooterIndex = footerRowStartIndex + sheet.footers.length;

          for (var ctr2 = footerRowStartIndex; ctr2 < maxFooterIndex; ctr2++) {

            const counter = ctr2 - footerRowStartIndex;

            const rowIndex = ctr2 + 1;

            const footerList = sheet.footers[counter].footerNames;

            const row = worksheet.getRow(rowIndex);

            row.height = sheet.footers[counter].height

            let colIndex = 1;
            // ITERATE NAMES AND SET VALUES IN EACH CELL
            for (var ctr4 = 0; ctr4 < footerList.length; ctr4++) {

              const cell = row.getCell(colIndex);

              if (footerList[ctr4].colspan > 1) {

                let currentCell = this[getColumnCode](colIndex, rowIndex);

                let cellSpan = this[getColumnCode]((footerList[ctr4].colspan - 1) + colIndex, rowIndex);

                worksheet.mergeCells(currentCell + ":" + cellSpan);
              }

              cell.value = footerList[ctr4].title;

              const style = footerList[ctr4].style;

              for (const [key, value] of Object.entries(style)) {

                cell[key] = value;
              }

              colIndex++;

              colIndex += footerList[ctr4].colspan - 1;
            }
          }
        }
      }


      return await workbook.xlsx.write(stream)
        .then(() => {
          return s3.upload({
            Key: filename,
            Bucket: bucket,
            Body: stream,
            ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          }).promise();
        });
    }

    createSheet = () => {
      const id = this.sheets.length + 1;
      const sheet = new Sheet(id);
      this.sheets.push(sheet);
      return sheet;
    }

    setExcelConfig = (sheets) => {
      if (!sheets) throw new Error('The "sheets" field is required.');
      if (sheets.length > 0) {
        for (let sheetCount = 0; sheetCount < sheets.length; sheetCount++) {
          const sheetElement = sheets[sheetCount];
          const sheet = this.createSheet();
          if (sheetElement.title)
            sheet.title = sheetElement.title;
          if (sheetElement.option)
            sheet.option = sheetElement.option;

          // Create all the headers
          for (let headerCount = 0; headerCount < sheetElement.headers.length; headerCount++) {
            const headerElement = sheetElement.headers[headerCount];
            if (headerElement.main === undefined || headerElement.main === null) throw new Error('Header "main" is required.');
            const header = sheet.createHeader(headerElement.main);
            if (headerElement.height)
              header.height = headerElement.height;

            // Add the columnss
            for (let headerNameCount = 0; headerNameCount < headerElement.headerNames.length; headerNameCount++) {
              const headerNameElement = headerElement.headerNames[headerNameCount];
              if (headerNameElement.title === undefined) throw new Error('Header "title" is required.');
              const headerName = header.createHeaderName(headerNameElement.title);
              if (headerNameElement.width)
                headerName.width = headerNameElement.width;
              if (headerNameElement.colspan)
                headerName.colspan = headerNameElement.colspan;
              if (headerNameElement.style)
                headerName.style = headerNameElement.style;
            }
          }

          // Add the rows
          if (sheetElement.rowData)
            sheet.rowData = sheetElement.rowData;
          if (sheetElement.headerFooter)
            sheet.headerFooter = sheetElement.headerFooter;

          if (sheetElement.footers !== undefined) {
            // Create all the footers
            for (let footerCount = 0; footerCount < sheetElement.footers.length; footerCount++) {
              const footerElement = sheetElement.footers[footerCount];
              const footer = sheet.createFooter();
              if (footerElement.height)
                footer.height = footerElement.height;

              // Add the footers
              for (let footerNameCount = 0; footerNameCount < footerElement.footerNames.length; footerNameCount++) {
                const footerNameElement = footerElement.footerNames[footerNameCount];
                if (footerNameElement.title === undefined) throw new Error('Footer "title" is required.');
                const footerName = footer.createFooterName(footerNameElement.title);
                if (footerNameElement.width)
                  footerName.width = footerNameElement.width;
                if (footerNameElement.colspan)
                  footerName.colspan = footerNameElement.colspan;
                if (footerNameElement.style)
                  footerName.style = footerNameElement.style;
              }
            }
          }
        }
      }
    }

    getSheetById = (id) => {
      for (let index = 0; index < this.sheets.length; index++) {
        const element = this.sheets[index];
        if (id === element.id) return element;
      }
      return null;
    }

    getSheetByTitle = (title) => {
      for (let index = 0; index < this.sheets.length; index++) {
        const element = this.sheets[index];
        if (title === element.title) return element;
      }
      return null;
    }
  }

  class Sheet {
    constructor(id) {
      this.id = id;
      this.title = "Sheet 1";
      this.option = { properties: { tabColor: { argb: '00ff00' } }, views: [{ showGridLines: false }] };
      this.headers = [];
      this.rowData = new RowData();
      this.footers = [];
      this.headerFooter = {};
    }
    createHeader = (isMain) => {
      for (var ctr = 0; ctr < this.headers.length; ctr++) {
        if (this.headers[ctr].main)
          throw new Error(
            'Sheet cannot contains multiple main header!'
          );
      }
      const id = this.headers.length + 1;
      const header = new Header(id, isMain);
      this.headers.push(header);
      return header;
    }

    createFooter = () => {
      const id = this.footers.length + 1;
      const footer = new Footer(id);
      this.footers.push(footer);
      return footer;
    }

    getHeaderById = (id) => {
      for (let index = 0; index < this.headers.length; index++) {
        const element = this.headers[index];
        if (id === element.id) return element;
      }
      return null;
    }

    getMainHeader = () => {
      for (let index = 0; index < this.headers.length; index++) {
        const element = this.headers[index];
        if (element.main) return element;
      }
      return null;
    }

    getFooterById = (id) => {
      for (let index = 0; index < this.footers.length; index++) {
        const element = this.footers[index];
        if (id === element.id) return element;
      }
      return null;
    }
  }

  class Header {
    constructor(id, isMain) {
      this.id = id;
      this.main = isMain;
      this.height = 20.5;
      this.headerNames = [];
    }
    createHeaderName = (title) => {
      const key = ["header", this.headerNames.length + 1].join('');
      const headerName = new HeaderName(key, title);
      this.headerNames.push(headerName);
      return headerName;
    }

    getHeaderNameByKey = (key) => {
      for (let index = 0; index < this.headerNames.length; index++) {
        const element = this.headerNames[index];
        if (key === element.key) return element;
      }
      return null;
    }

    getHeaderNameByTitle = (title) => {
      for (let index = 0; index < this.headerNames.length; index++) {
        const element = this.headerNames[index];
        if (title === element.title) return element;
      }
      return null;
    }

  }

  class HeaderName {
    constructor(key, title) {
      this.key = key;
      this.title = title;
      this.width = 20;
      this.colspan = 1;
      this.style = new Style();
    }
  }

  class Footer {
    constructor(id) {
      this.id = id;
      this.height = 20.5;
      this.footerNames = [];
    }
    createFooterName = (title) => {
      const key = ["footer", this.footerNames.length + 1].join('');
      const footerName = new FooterName(key, title);
      this.footerNames.push(footerName);
      return footerName;
    }

    getFooterNameByKey = (key) => {
      for (let index = 0; index < this.footerNames.length; index++) {
        const element = this.footerNames[index];
        if (key === element.key) return element;
      }
      return null;
    }

    getFooterNameByTitle = (title) => {
      for (let index = 0; index < this.footerNames.length; index++) {
        const element = this.footerNames[index];
        if (title === element.title) return element;
      }
      return null;
    }

  }

  class FooterName {
    constructor(key, title) {
      this.key = key;
      this.title = title;
      this.width = 20;
      this.colspan = 1;
      this.style = new Style();
    }
  }

  class Style {
    constructor() {
      this.font = { bold: false, color: { argb: '000000' }, name: 'Arial', size: 9 };
      this.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
      this.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      this.alignment = { horizontal: 'left' };
    }
  }

  class RowData {
    constructor() {
      this.style = new Style();
      this.height = 45.2;
      this.rows = [];
    }
  }

  return Excel2S3;
})();


module.exports = Excel2S3;
