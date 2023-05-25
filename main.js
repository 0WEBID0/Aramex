
// <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
      
      function time(){
        document.getElementById("sendtable").style.display = 'none' ;
        document.getElementById("Reloaded").style.display = 'block ' ;
        document.getElementById("dowaramex0").style.display = 'block ' ;
        document.getElementById("dowaramex1").style.display = 'block ' ;

      }

      async function loadPDF() {
        
  const fileInput = document.getElementById('servicestimeA');
  const file = fileInput.files[0];
  
  if (file) {
    const arrayBuffer = await file.arrayBuffer();
    const uint8Array = new Uint8Array(arrayBuffer);
    
    const loadingTask = pdfjsLib.getDocument(uint8Array);
    const pdf = await loadingTask.promise;
  
    // استخدم الملف PDF كما هو مطلوب
    const pdfContent = await extractPDFText(pdf);
    // ...
  }
}
  
      async function extractPDFText(pdf) {
        const xlsxData = await loadXLSX();
        const xlsxDataD = await loadXLSXD();
        const {
    column13,
    column23,
    column33,
    column43,
    column53,
    column63
  } = xlsxDataD;

    const {
      column22,
    column32,
    column42, //kilo
    column52
  } = xlsxData;

  // استخدام المتغيرات column12 و column22 و column32 و column42 و column52 هنا في extractPDFText(pdf)
  // ...
  // مثال على كيفية استخدام المتغيرات في extractPDFText(pdf)

  const column1 = [];
    const column2 = [];
    const column3 = [];
    const column4 = [];
    const column5 = [];
    const column6 = [];
    const column7 = [];
    const column8 = [];
    const column9 = [];
    const column10 = [];
      const countries = [
'India',
'Qatar',
'Saudi Arabia',
'Bahrain',
'Oman',
'Utd.Arab',
'United',
'Yemen',
'Denmark',
'Morocco',
'Jordan',
'Australia',
'USA',
'Sri Lanka',
'Algeria',
'Iraq',
'Egypt',
'Palestine',
'Lebanon',
'Tunisia',
'Mauritania',
'Sudan',
'UAE',
'Tunis',
'Afghanistan',
'China',
'Canada',
'Brazil',
'Türkiye',
'Poland',
'Russia',
'Holland',
'Portugal',
'Spain',
'Congo',
'Argentina',
'Nepal',
'Bahrain',

'GDX',
'GPX',
'LHS',
'BIO',
'DPX',
'VPX',
'PDX',
'EXP',
'LAND',
'EDX',
'EPX',
'DGX',
'DGG',
'PPX',
'PXP',
'PLX',
'RTC',
// قائمة الدول الأخرى
];
    let totalnet = 0 ;
    let totalwight = 0 ;
    let totalpriceE = 0 ;
    let totalWodex = 0 ;
    let totalsale = 0 ;
    let totwightwodex = 0 ;
    let startReading = false;
    let lineCounter = 0;
    let currentLine = '';
    var table = $('#example tbody');
    
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
      const page = await pdf.getPage(pageNumber);
      const pageText = await page.getTextContent();
  
      for (let item of pageText.items) {
        const { str } = item;
        if (str.includes('S.No')) {
          startReading = true;
        }
  
        if (
          startReading &&
          (!isEnglishAlphabet(str.charAt(0)) || countries.includes(str.trim())) &&
          str.trim() !== '' &&
          !str.includes(',') &&
          !str.includes('•') &&
          !str.includes('of') &&
          !str.includes('TO') &&
          !str.includes('ON')
        ) {
          if (lineCounter < 9) {
            currentLine += str.trim() + ' ';
            lineCounter++;
          } else {
            currentLine += str.trim();
            const columnsInRow = currentLine.split(' ');
            column1.push(parseFloat(columnsInRow[0])); //number
            column2.push(parseFloat(columnsInRow[1])); //id
            column3.push(columnsInRow[2]); //date
            column4.push(parseFloat(columnsInRow[3]).toFixed(3));//net amount
            column5.push(parseFloat(columnsInRow[4]));
            column6.push(parseFloat(columnsInRow[5]));
            columnsInRow[7] = Math.ceil(parseFloat(columnsInRow[7]) * 2) / 2 ;
            column7.push(columnsInRow[7]); //wight
            column8.push(columnsInRow[8]); //KG
            column9.push(columnsInRow[9]); //prof type
            column10.push(columnsInRow[10]); // Distantion
            // totalnet += parseFloat(columnsInRow[3]);
            // totalwight += parseFloat(columnsInRow[7]);
            currentLine = '';
            lineCounter = 0;
          }
        }
      }
    }
    
    let column1_1 = [] ; //profit
    let column1_2 = [] ; //less
    let column1_3 = [] ; //price change
    let count = [] ;
    let count21 = []
    let count31 = [] ;
    let count61 = []
    let number = 0
    let totalprofit = 0.0 ;
    let totalless = 0.0 ;
      for(let l = 0 ; l < (column32.length) ; l++){
              for(let r = 0 ; r < (column1.length) ; r++){
                if(column32[l] == column2[r]){
                  // console.log(column32[l] , column2[r])
                  totalwight += parseFloat(column7[r]);
                  totalnet += (parseFloat(column4[r])) ;
                  totalpriceE += (column52[l]);
                  totalsale += parseFloat(((column52[l] - (parseFloat(column4[r]))).toFixed(3)));
                  count[number] = (column1[r -1]) 
                  table.append('<tr style="height: 83px;" ><td class="u-table-cell" >'+column1[r]+'</td><td class="u-table-cell" >'+column2[r]+'</td><td class="u-table-cell">'+column22[l]+'</td><td class="u-table-cell" >'+column3[r]+'</td><td class="u-table-cell" >'+column10[r]+'</td><td class="u-table-cell">'+column9[r]+'</td><td class="u-table-cell" >'+(column7[r]+ " KG")+'</td><td class="u-table-cell" id="wightWodex'+[r+1]+'" ></td><td class="u-table-cell green" id="profit'+[r+1]+'" ></td><td class="u-table-cell red" id="less'+[r+1]+'" ></td><td class="u-table-cell">'+column4[r]+'</td><td class="u-table-cell" id="change'+[r+1]+'" ></td><td class="u-table-cell" id="priceWODEX'+[r+1]+'"></td><td class="u-table-cell">'+(column52[l])+'</td><td class="u-table-cell" id="profitsale'+[r+1]+'">'+((column52[l] - (parseFloat(column4[r]))).toFixed(3))+'</td></tr>');
                  count21[number] = column42[l] ;
                  count31[number] = column52[l] ;
                  count61[number] = column22[l];
                  number += 1 ;
                  if(column42[l] > column7[r]){
                    totwightwodex += (parseFloat(column42[l]))
                    document.getElementById("wightWodex"+[r+1]).innerHTML = column42[l]+ " KG"  ;
                    let profitwight = (column42[l] - column7[r]) ;
                    document.getElementById("profit"+[r+1]).innerHTML = profitwight + " KG" ;
                    totalprofit += profitwight ;
                    column1_1 += profitwight ;
                    column1_2 += " " ;
                  }else{
                    document.getElementById("wightWodex"+[r+1]).innerHTML = column42[l]+ " KG"  ;
                    let lesswight = (column7[r] - column42[l]) ;
                    document.getElementById("less"+[r+1]).innerHTML = lesswight + " KG" ;
                    totalless += lesswight
                    column1_2 += lesswight ;
                    column1_1 += " " ;
                  }
            }
          }
          }


number = 0 ;
let count41 = [] ;
let count51 = []

for(let l = 0 ; l < (column1.length) ; l++){
  for(let r = 0 ; r < (column32.length) ; r++){
    if(column2[l] == column32[r]){
                        if((((parseFloat(column52[r])) - (parseFloat(column4[l]))).toFixed(3)) > 0){
                        document.getElementById("profitsale"+(l+1)).style.color = "green";}
                        else{
                          document.getElementById("profitsale"+(l+1)).style.color = "red";
                        }
      for(let s = 0 ; s < (column13.length) ; s++){
  if (["Saudi", "Bahrain", "Oman", "Qatar", "Utd.Arab"].includes(column10[l]) ) {
      if(column7[l] == column13[s]){
        totalWodex += (parseFloat(column33[s]))  ;
        count41[number] =(parseFloat( column33[s])) ;
        number += 1 ;
              document.getElementById("priceWODEX"+[l+1]).innerHTML = column33[s] ;
                      if(column4[l] != column33[s]){
                      document.getElementById("change"+[l+1]).innerHTML = "Change" ;
                      }
                    }
              }else if (["United"].includes(column10[l])){
                if(column7[l] == column43[s]){
                  totalWodex += (parseFloat(column33[s]))  ;
                  count51[number] =(parseFloat( column63[s])) ;
                  document.getElementById("priceWODEX"+[l+1]).innerHTML = column63[s] ;
                  number += 1 ;
                if(column4[l] != column63[s]){
                  document.getElementById("change"+[l+1]).innerHTML = "Change" ;
                        if(column52[l] > column63[s]){
                        document.getElementById("profitsale"+(l+1)).style.color = "green";}
                        else{
                          document.getElementById("profitsale"+(l+1)).style.color = "red";
                        }
                      }
                    }
                  }
                    
                    
                  
                }
                    }

  }
}

// console.log(count51)

        document.getElementById("totalprofit").innerHTML = totalprofit + " KG" ;
        document.getElementById("totalless").innerHTML = totalless + " KG" ;
        document.getElementById("totwightwodex").innerHTML =  parseFloat((totwightwodex))+ ' KG' ;
        document.getElementById("totalWodex").innerHTML =  parseFloat((totalWodex).toFixed(3)) ;
        document.getElementById("totalpriceE").innerHTML =  parseFloat((totalpriceE).toFixed(3));
        document.getElementById("totalsale").innerHTML =  parseFloat((totalsale).toFixed(3)) ;
    document.getElementById("totnet").innerHTML =  totalnet.toFixed(3) ;
    document.getElementById("totwight").innerHTML =  totalwight + ' KG' ;
    // console.log(column33 , column63)
    document.getElementById('dowaramex0').addEventListener('click', function() {
      downloadAramex(column1,column2,column3,column10,column9,column7,column1_1 ,column1_2 ,column4 ,column1_3 , column42 , column33 , column63 ,column52 ,count ,number ,count21 ,count31 ,count41 ,count51 ,totalWodex ,totalless ,totalnet ,totalprofit ,totalsale ,totalwight ,totalpriceE ,totwightwodex ,count61);
  });
  }
  
  function isEnglishAlphabet(char) {
    const charCode = char.charCodeAt(0);
    return (charCode >= 65 && charCode <= 90) || (charCode >= 97 && charCode <= 122); // رموز A إلى Z و a إلى z في الأبجدية الإنجليزية
  }
      // استدعاء الدالة عند تحميل الصفحة
      


      function generateRandomData(column1,column2,column3,column10,column9,column7,column1_1 ,column1_2 ,column4 ,column1_3 , column42 , column33 , column63 ,column52 ,count ,number ,count21 ,count31 ,count41 ,count51 ,totalWodex ,totalless ,totalnet ,totalprofit ,totalsale ,totalwight ,totalpriceE ,totwightwodex ,count61) {
            var data = [];
            console.log(column7[count[69]])
            for (var i = -1; i < (number+1) ; i++) {
                var row = [];
                for (var j = 0; j < 15; j++) {
                  if (i==-1){
                    if(j==0)
                    row.push("S.NO") ;
                    else if (j==1)
                    row.push("HAWB");
                    else if (j==2)
                    row.push("Invoice NO.");
                    else if (j==3)
                    row.push("Pack up date");
                    else if (j==4)
                    row.push("country");
                    else if (j==5)
                    row.push("prof type");
                    else if (j==6)
                    row.push("Weight Aremax");
                    else if (j==7)
                    row.push("wight WODEX");
                    else if (j==8)
                    row.push("Profit");
                    else if (j==9)
                    row.push("Less");
                    else if (j==10)
                    row.push("Net Amount Aremax");
                    else if (j==11)
                    row.push("Price Change");
                    else if (j==12)
                    row.push("Price Wodex");
                    else if (j==13)
                    row.push("Selling price with Aramex");
                    else if (j==14)
                    row.push("profit sale");
                    // document.getElementById(" " + count[i]).innerHTML = 
                  }else if( i < number){
                    if(j==0)
                    row.push(column1[count[i]]);
                    else if (j==1)
                    row.push(column2[count[i]]);
                    else if (j==2)
                    row.push(count61[i]);
                    else if (j==3)
                    row.push(column3[count[i]]);
                    else if (j==4)
                    row.push(column10[count[i]]);
                    else if (j==5)
                    row.push(column9[count[i]]);
                    else if (j==6)
                    row.push(column7[count[i]] + " KG");
                    else if (j==7)
                    row.push(count21[i] + " KG");
                    else if (j==8){
                      if(count21[i] > column7[count[i]]){
                    row.push((count21[i] - column7[count[i]])+ " KG");
                    }else{
                      row.push(" ");
                    }
                    }
                    else if (j==9){
                      if(count21[i] <= column7[count[i]]){
                    row.push((count21[i] - column7[count[i]])+ " KG");
                    }else{
                      row.push(" ");
                    }
                    }
                    else if (j==10)
                      row.push(column4[count[i]]);
                    else if (j==11){
                      if (["Saudi", "Bahrain", "Oman", "Qatar", "Utd.Arab"].includes(column10[i]) ) {
                        if(count41[i] != column4[count[i]]){
                          row.push("change");
                        }else{
                          row.push(" ");
                        }
                      }else if (["United"].includes(column10[i])){
                        if(count51[i] != column4[count[i]]){
                          row.push("change");
                        }else{
                          row.push(" ");
                        }
                      }else{
                        row.push(" ");
                      }
                    }
                    else if (j==12){
                      if (["Saudi", "Bahrain", "Oman", "Qatar", "Utd.Arab"].includes(column10[i]) ) {
                        row.push(count41[i]);
                      }else if (["United"].includes(column10[i])){
                        row.push(count51[i]);
                      }else{
                        row.push(" ");
                      }
                    } // or column63
                    else if (j==13)
                    row.push(count31[i]);
                    else if (j==14)
                    row.push((count31[i]) - (column4[count[i]]));
                  }else if (i == number){
                    if (j==0)
                    row.push("TOTAL");
                    else if (j==1)
                    row.push(" ");
                    else if (j==2)
                    row.push(" ");
                    else if (j==3)
                    row.push(" ");
                    else if (j==4)
                    row.push(" ");
                    else if (j==5)
                    row.push(" ");
                    else if (j==6)
                    row.push(totalwight +" KG");
                    else if (j==7)
                    row.push(totwightwodex +" KG");
                    else if (j==8)
                    row.push(totalprofit +" KG");
                    else if (j==9)
                    row.push(totalless +" KG");
                    else if (j==10)
                    row.push(totalnet);
                    else if (j==11)
                    row.push("Done");
                    else if (j==12)
                    row.push(totalWodex);
                    else if (j==13)
                    row.push(totalpriceE);
                    else if (j==14)
                    row.push(totalsale);
                }
              }
                data.push(row);
            }
            return data;
        }

        function downloadAramex(column1,column2,column3,column10,column9,column7,column1_1 ,column1_2 ,column4 ,column1_3 , column42 , column33 , column63 ,column52 ,count ,number ,count21 ,count31 ,count41 ,count51 ,totalWodex ,totalless ,totalnet ,totalprofit ,totalsale ,totalwight ,totalpriceE ,totwightwodex ,count61) {
  var fileName = "Aramex Invoice.xlsx";

  var workbook = XlsxPopulate.fromBlankAsync()
    .then(function (workbook) {
      var sheet = workbook.sheet(0);
      sheet.name("Sheet1");

      var data = generateRandomData(column1,column2,column3,column10,column9,column7,column1_1 ,column1_2 ,column4 ,column1_3 , column42 , column33 , column63 ,column52 ,count ,number ,count21 ,count31 ,count41 ,count51 ,totalWodex ,totalless ,totalnet ,totalprofit ,totalsale ,totalwight ,totalpriceE ,totwightwodex ,count61);

      for (var i = 0; i < data.length; i++) {
        var rowData = data[i];
        for (var j = 0; j < rowData.length; j++) {
          var cell = sheet.cell(i + 1, j + 1);
          cell.style({ border: true});
          cell.value(rowData[j]);
        }
      }

      var usedRange = sheet.usedRange();
      var startCell = usedRange.startCell();
      var endCell = usedRange.endCell();
      var columnCount = endCell.columnNumber() - startCell.columnNumber() + 1;

      // تعيين حجم الأعمدة
      for (var colIndex = startCell.columnNumber(); colIndex <= endCell.columnNumber(); colIndex++) {
        sheet.column(colIndex).width(11); // تعيين عرض العمود
      }


      return workbook.outputAsync();
    })
    .then(function (blob) {
      var link = document.createElement('a');
      link.href = window.URL.createObjectURL(blob);
      link.download = fileName;

      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    });
}

// <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
    





// <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
      

   async function loadXLSX() {
  const fileInput = document.getElementById('servicestimeB');
  const file = fileInput.files[0];

  // تحميل محتوى الملف XLSX
  const reader = new FileReader();
  return new Promise((resolve, reject) => {
    reader.onload = function (event) {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      // تخزين محتوى الملف في صفحة الويب
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      const column12 = [];
      const column22 = [];
      const column32 = [];
      const column42 = [];
      const column52 = [];

      // الحصول على حقول العمود المطلوبة (مثلاً العمود A و B و C و D و E)
      const desiredColumns = ['A', 'B', 'C', 'D', 'E'];

      const range = XLSX.utils.decode_range(worksheet['!ref']);
      for (let i = 0; i < (desiredColumns.length); i++) {
        const column = desiredColumns[i];

        for (let row = range.s.r + 4; row <= (range.e.r); row++) {
          const cellAddress = column + row;
          let cellValue = worksheet[cellAddress]?.v;

          // تحويل القيمة إلى نص في حالة العمود الأول
          if (i === 0) {
            const cellDate = XLSX.SSF.parse_date_code(cellValue);
            cellValue = `${cellDate.m}/${cellDate.d}/${cellDate.y}`;
          }

          // تخزين قيم العمود في المصفوفات المنفصلة
          switch (i) {
            case 0:
              column12.push(cellValue);
              break;
            case 1:
              column22.push(cellValue);
              break;
            case 2:
              column32.push(cellValue);
              break;
            case 3:
              column42.push(cellValue);
              break;
            case 4:
              column52.push(cellValue);
              break;
          }
        }
      }

      resolve({
      column22,
        column32,
        column42,
        column52
      });
    };

    reader.onerror = function (event) {
      reject(event.target.error);
    };

    reader.readAsArrayBuffer(file);
  });
}


  

  // <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
    





// <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
  
async function loadXLSXD() {
const fileInput = document.getElementById('servicestimeD');
const file = fileInput.files[0];

// تحميل محتوى الملف XLSX
const reader = new FileReader();
return new Promise((resolve, reject) => {
reader.onload = async function(event) {
  const data = new Uint8Array(event.target.result);
  const workbook = XLSX.read(data, { type: 'array' });

  // قراءة جميع الورقات
  const sheetNames = workbook.SheetNames;

  let column13 = [];
  let column23 = [];
  let column33 = [];
  let column43 = [];
  let column53 = [];
  let column63 = [];

  sheetNames.forEach((sheetName, index) => {
    const worksheet = workbook.Sheets[sheetName];

    // الحصول على حقول العمود المطلوبة (مثلاً العمود A و B و C و D و E)
    const desiredColumns = ['A', 'B', 'C', 'D'];

    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let i = 0; i < desiredColumns.length; i++) {
      const column = desiredColumns[i];

      for (let row = range.s.r + 4; row <= range.e.r; row++) {
        const cellAddress = column + row;
        let cellValue = worksheet[cellAddress]?.v;

        // تخزين قيم العمود في المصفوفات المنفصلة
        switch (index) {
          case 0:
            if (i === 0) column13.push(cellValue);
            if (i === 1) column23.push(cellValue);
            if (i === 2) column33.push(cellValue.toFixed(3));
            break;
          case 1:
            if (i === 0) column43.push(cellValue);
            if (i === 1) column53.push(cellValue);
            if (i === 2) column63.push(cellValue.toFixed(3));
            break;
          // قم بإضافة حالات إضافية للورقات الأخرى حسب الحاجة
        }
      }
    }
  });
  resolve({
    column13,
    column23,
    column33,
    column43,
    column53,
    column63
  });
};
reader.onerror = function (event) {
  reject(event.target.error);
};

reader.readAsArrayBuffer(file);
});
}


// <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
  






// <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
  
async function loadXLSXE() {
const fileInput = document.getElementById('servicestimeE');
const file = fileInput.files[0];

// تحميل محتوى الملف XLSX
const reader = new FileReader();
return new Promise((resolve, reject) => {
  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    // تخزين محتوى الملف في صفحة الويب
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    const column12 = [];
    const column24 = [];
    const column34 = [];
    const column44 = [];
    const column54 = [];
    const column64 = [];
    const column74 = [];
    const column84 = [];

    // الحصول على حقول العمود المطلوبة (مثلاً العمود A و B و C و D و E)
    const desiredColumns = ['A', 'B', 'C', 'D', 'E' ,'F' ,'G' ,'H'];

    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let i = 0; i < (desiredColumns.length); i++) {
      const column = desiredColumns[i];

      for (let row = range.s.r + 4; row <= (range.e.r); row++) {
        const cellAddress = column + row;
        let cellValue = worksheet[cellAddress]?.v;

        // تحويل القيمة إلى نص في حالة العمود الأول
        if (i === 0) {
          const cellDate = XLSX.SSF.parse_date_code(cellValue);
          cellValue = `${cellDate.m}/${cellDate.d}/${cellDate.y}`;
        }

        // تخزين قيم العمود في المصفوفات المنفصلة
        switch (i) {
          case 0:
            column12.push(cellValue);
            break;
          case 1:
            column24.push(cellValue);
            break;
          case 2:
            column34.push(cellValue);
            break;
          case 3:
            column44.push(cellValue);
            break;
          case 4:
            column54.push(cellValue);
            break;
            case 5:
            column64.push(cellValue);
            break;
            case 6:
            column74.push(cellValue);
            break;
            case 7:
            column84.push(cellValue);
            break;
        }
      }
    }
// console.log(column84)
    resolve({
      column24,
      column34,
      column44,
      column54,
      column84
    });
  };

  reader.onerror = function (event) {
    reject(event.target.error);
  };

  reader.readAsArrayBuffer(file);
});
}


// <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
  


async function loadPDFB() {
      
      const fileInput = document.getElementById('servicestimeC');
      const file = fileInput.files[0];
      
      if (file) {
        const arrayBuffer = await file.arrayBuffer();
        const uint8Array = new Uint8Array(arrayBuffer);
        
        const loadingTask = pdfjsLib.getDocument(uint8Array);
        const pdf = await loadingTask.promise;
      
        // استخدم الملف PDF كما هو مطلوب
        const pdfContent = await extractPDFTextB(pdf);
        // ...
      }
    }

    async function extractPDFTextB(pdf) {
      const xlsxData = await loadXLSXE();
  const {
  column24,
  column34,
  column44, //kilo
  column54,
  column84
} = xlsxData;
// console.log(column34 ,column44 ,column54 )
let totalwightWodexB = 0 ;
let totalwightB = 0 ;
let totalpeiceWodexB = 0 ;
let totalpriceB = 0 ;
let totalprofitB = 0 ;
let totalless = 0 ;
let totwightwodexprofit = 0 ;
let totwightwodexless = 0 ;

const column1 = [];
const column2 = [];
const column3 = [];
const column4 = [];
const column5 = [];
const column6 = [];
const column7 = [];
const column8 = [];
const column9 = [];
const column10 = [];
const column11 = [];
const column12 = [];
const column13 = [];
let countB = []
let count1B = []
let count2B = []
let countprofitwight = [] ;
let countlesswight = []
let countpricewodex = []
const pdfContent = [];
let numberB = 0 ;
var table = $('#exampleB tbody');
let startReading = false;
let lineCounter = 0;
let currentLine = '';
for (let pageNumber = 2; pageNumber <= pdf.numPages; pageNumber++) {
  const page = await pdf.getPage(pageNumber);
  const pageText = await page.getTextContent();
  for (let item of pageText.items) {
    const { str } = item;
    if (str.trim().length >= 6) {
      startReading = true;
    }
    if (
      startReading &&
      (!isEnglishAlphabet(str.charAt(0))) &&
      str.trim() !== '' &&
      !str.includes(',') &&
      !str.includes('•') &&
      !str.includes('of') &&
      !str.includes('TO') &&
      !str.includes('ON')
    ) {
      if (lineCounter < 11) {
        currentLine += str.trim() + ' ';
        lineCounter++;
      } else {
        currentLine += str.trim();
        const columnsInRow = currentLine.split(' ');
        column1.push(parseFloat(columnsInRow[0])); //id
        column2.push(parseFloat(columnsInRow[1])); //snipper
        column3.push(columnsInRow[2]); //date
        column4.push(columnsInRow[3]); //wight
        column5.push(columnsInRow[4]);
        column6.push(parseFloat(columnsInRow[5]));
        column7.push(columnsInRow[6]); 
        column8.push(columnsInRow[7]);
        column9.push(columnsInRow[8]);
        column10.push(columnsInRow[9]); 
        column11.push(columnsInRow[10]);
        column12.push(columnsInRow[11]);
        column13.push(columnsInRow[12]); //net amount
        pdfContent.push(currentLine);
        currentLine = '';
        lineCounter = 0;
        }
      }
    }
  for(let l = 0 ; l < (column34.length) ; l++){
            for(let r = 0 ; r < (column1.length) ; r++){
              if(column34[l] === column1[r]){
                countB[numberB] = r ;
                count1B[numberB] = column24[l]
                count2B[numberB] = column84[l]
                let profitwight = 0 ;

                countpricewodex[numberB] = column54[l] ;
                if (column44[l] > column4[r]) {
                  profitwight = (column44[l] - column4[r]) ;
                  totwightwodexprofit += (profitwight)
                  countprofitwight[numberB] = profitwight
                  countlesswight[numberB] = " "
  table.append('<tr style="height: 83px;"><td class="u-table-cell">'+(numberB +1)+'</td><td class="u-table-cell">' + column1[r] + '</td><td class="u-table-cell">'+column24[l]+'</td><td class="u-table-cell">' + column3[r] + '</td><td class="u-table-cell">'+column84[l]+'</td><td class="u-table-cell">' + (column4[r] + " KG") + '</td><td class="u-table-cell">' + (column44[l] + " KG") + '</td><td class="u-table-cell green">' + profitwight+" KG" + '</td><td class="u-table-cell"></td><td class="u-table-cell">' + column13[r] + '</td><td class="u-table-cell">' + column54[l] + '</td><td class="u-table-cell id="'+"profchaB"+[r]+'">' + ((column54[l] - column13[r]).toFixed(3)) + '</td></tr>');
} else {
lesswight = ( column4[r] - column44[l]) ;
totwightwodexless += (lesswight)
countprofitwight[numberB] = " "
countlesswight[numberB] = lesswight
  table.append('<tr style="height: 83px;"><td class="u-table-cell">'+(numberB +1)+'</td><td class="u-table-cell">' + column1[r] + '</td><td class="u-table-cell">'+column24[l]+'</td><td class="u-table-cell">' + column3[r] + '</td><td class="u-table-cell">'+column84[l]+'</td><td class="u-table-cell">' + (column4[r] + " KG") + '</td><td class="u-table-cell">' + (column44[l] + " KG") + '</td><td class="u-table-cell "></td><td class="u-table-cell red">' + lesswight+" KG" + '</td><td class="u-table-cell">' + column13[r] + '</td><td class="u-table-cell">' + column54[l] + '</td><td class="u-table-cell id="'+"profchaB"+[r]+'"">' + ((column54[l] - column13[r]).toFixed(3)) + '</td></tr>');
}
                // number += 1 ;
                // if(column44[l] > column4[r]){
                  totalwightWodexB += (parseFloat(column44[l]))
                  totalwightB += (parseFloat(column4[r]))
                  totalpeiceWodexB += (parseFloat(column54[l]))
                  totalpriceB += (parseFloat(column13[r]))
                  totalprofitB += (parseFloat(((column54[l] - column13[r])).toFixed(3)))
                  numberB += 1

          }
        }
        }

if (currentLine !== '') {
  pdfContent.push(currentLine);
}
// console.log(countB)
document.getElementById("totwightwodexprofit").innerHTML = (parseFloat( totwightwodexprofit)).toFixed(1) +" KG"
document.getElementById("totwightwodexless").innerHTML = (parseFloat( totwightwodexless)).toFixed(1) +" KG"
document.getElementById("totalwightWodexB").innerHTML = (parseFloat( totalwightWodexB)).toFixed(1) +" KG"
document.getElementById("totalwightB").innerHTML = (parseFloat( totalwightB)).toFixed(1) +" KG"
document.getElementById("totalpeiceWodexB").innerHTML = (parseFloat( totalpeiceWodexB)).toFixed(3)
document.getElementById("totalpriceB").innerHTML = (parseFloat( totalpriceB)).toFixed(3)
document.getElementById("totalprofitB").innerHTML = (parseFloat( totalprofitB)).toFixed(3)

document.getElementById('dowaramex1').addEventListener('click', function() {
    downloadAramexB(totwightwodexprofit ,totalwightB ,totalwightWodexB ,totalpeiceWodexB ,totalpriceB ,totalprofitB ,column1 ,column2 ,column3 ,column44 ,column4 ,column54 ,column13 ,countB ,numberB ,countprofitwight ,countpricewodex ,count1B ,count2B ,totwightwodexless ,countlesswight);
});

return pdfContent.join('\n');
}
    }

function isEnglishAlphabet(char) {
const charCode = char.charCodeAt(0);
return (charCode >= 65 && charCode <= 90) || (charCode >= 97 && charCode <= 122);  // رموز A إلى Z في الأبجدية الإنجليزية
}




    // <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->
// <------------------------------------------------------------------------------------------------------------------------------->



function generateRandomDataB(totwightwodexprofit ,totalwightB ,totalwightWodexB ,totalpeiceWodexB ,totalpriceB ,totalprofitB ,column1 ,column2 ,column3 ,column44 ,column4 ,column54 ,column13 ,countB ,numberB ,countprofitwight ,countpricewodex ,count1B ,count2B ,totwightwodexless ,countlesswight) {
          var data = [];
          // console.log(count21)
          for (var i = -1; i < (numberB+1) ; i++) {
              var row = [];
              for (var j = 0; j < 11; j++) {
                if (i==-1){
                  if(j==0)
                  row.push("S.NO") ;
                  else if (j==1)
                  row.push("HAWB");
                  else if (j==2)
                  row.push("Invoice NO.");
                  else if (j==3)
                  row.push("date");
                  else if (j==4)
                  row.push("country");
                  else if (j==5)
                  row.push("wight DHL");
                  else if (j==6)
                  row.push("wight Wodex");
                  else if (j==7)
                  row.push("Profit wight");
                  else if (j==8)
                  row.push("Less wight");
                  else if (j==9)
                  row.push("Total DHL");
                  else if (j==10)
                  row.push("Price Wodex");
                  else if (j==11)
                  row.push("Profit Total");
                  // document.getElementById(" " + count[i]).innerHTML = 
                }else if( i < numberB){
                  if(j==0)
                  row.push(i+1);
                  else if(j==1)
                  row.push(column1[countB[i]]);
                  else if (j==2)
                  row.push(count1B[i]);
                  else if (j==3)
                  row.push(column3[countB[i]]);
                  else if (j==4)
                  row.push(count2B[i]);
                  else if (j==5)
                  row.push(column4[countB[i]]+ " KG");
                  else if (j==6)
                  row.push(column44[i]+ " KG");
                  else if (j==7){
                  if(countprofitwight[i] != " "){
                  row.push(countprofitwight[i]+ " KG");
                  }else{
                    row.push(" ");
                  }
                }
                  else if (j==8){
                  if(countlesswight[i] != " "){
                  row.push(countlesswight[i]+ " KG");
                  }else{
                    row.push(" ");
                  }
                }
                  else if (j==9)
                    row.push(column13[countB[i]]);
                  else if (j==10)
                  row.push(countpricewodex[i]);
                  else if (j==11)
                    row.push(((countpricewodex[i] - column13[countB[i]]).toFixed(2)));
                }else if (i == numberB){
                  if (j==0)
                  row.push("TOTAL");
                  else if (j==1)
                  row.push(" ");
                  else if (j==2)
                  row.push(" ");
                  else if (j==3)
                  row.push(" ");
                  else if (j==4)
                  row.push(" ");
                  else if (j==5)
                  row.push(totalwightB +" KG");
                  else if (j==6)
                  row.push(totalwightWodexB +" KG");
                  else if (j==7)
                  row.push(totwightwodexprofit +" KG");
                  else if (j==8)
                  row.push(totwightwodexless +" KG");
                  else if (j==9)
                  row.push(totalpriceB);
                  else if (j==10)
                  row.push(totalpeiceWodexB);
                  else if (j==11)
                  row.push(totalprofitB);
                }
            }
              data.push(row);
          }
          return data;
      }

      function downloadAramexB(totwightwodexprofit ,totalwightB ,totalwightWodexB ,totalpeiceWodexB ,totalpriceB ,totalprofitB ,column1 ,column2 ,column3 ,column44 ,column4 ,column54 ,column13 ,countB ,numberB ,countprofitwight ,countpricewodex ,count1B ,count2B ,totwightwodexless ,countlesswight) {
var fileName = "DHL Invoice.xlsx";

var workbook = XlsxPopulate.fromBlankAsync()
  .then(function (workbook) {
    var sheet = workbook.sheet(0);
    sheet.name("DHL");

    var data = generateRandomDataB(totwightwodexprofit ,totalwightB ,totalwightWodexB ,totalpeiceWodexB ,totalpriceB ,totalprofitB ,column1 ,column2 ,column3 ,column44 ,column4 ,column54 ,column13 ,countB ,numberB ,countprofitwight ,countpricewodex ,count1B ,count2B ,totwightwodexless ,countlesswight);

    for (var i = 0; i < data.length; i++) {
      var rowData = data[i];
      for (var j = 0; j < rowData.length; j++) {
        var cell = sheet.cell(i + 1, j + 1);
        cell.value(rowData[j]);
        cell.style({ border: true });
      }
    }

    var usedRange = sheet.usedRange();
    var startCell = usedRange.startCell();
    var endCell = usedRange.endCell();
    var columnCount = endCell.columnNumber() - startCell.columnNumber() + 1;

    // تعيين حجم الأعمدة
    for (var colIndex = startCell.columnNumber(); colIndex <= endCell.columnNumber(); colIndex++) {
      sheet.column(colIndex).width(11); // تعيين عرض العمود
    }

    return workbook.outputAsync();
  })
  .then(function (blob) {
    var link = document.createElement('a');
    link.href = window.URL.createObjectURL(blob);
    link.download = fileName;

    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  });
}

  
