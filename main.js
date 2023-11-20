function download() {
  var iframe = document.getElementById("output-iframe");
  if (iframe.src != "") {
    const anchor = document.createElement('a');
    anchor.href = iframe.src;
    anchor.download = "rozvrh.pdf";
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
  } else {
    alert("Nejdřive nechte vytvořit rozvrh");
  }
}

async function renderTimetable() {
  var iframe = document.getElementById("output-iframe");
  iframe.src = "";

  var inputElement = document.getElementById("file-input")
  if (inputElement.files.length < 1) {
    alert("Není vybrán žádný soubor");
    return;
  }

  var f = inputElement.files[0];
  var reader = new FileReader();
  reader.onerror = function(e) {
    alert("Došlo k chybě při otevírání souboru");
  };
  reader.onload = async function (e) {
    var data = e.target.result;

    // Create PDF doc
    const doc = new PDFDocument({ size: 'A4', layout: 'landscape' });
    const font = await fetch('RobotoCondensed-Regular.ttf')
    const arrayBuffer = await font.arrayBuffer()
    doc.registerFont('RobotoCondensed-Regular', arrayBuffer)

    var stream = doc.pipe(blobStream());

    var workbook = XLSX.read(data, { type: 'binary' });
    workbook.SheetNames.forEach(async (sheetName, index) => {
      if (index > 0) {
        doc.addPage({ size: 'A4', layout: 'landscape' });
      }
      var jsa = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "", raw: false });
      console.log(jsa);
      await renderSheet(doc, jsa, sheetName);
    });

    // Finalize PDF file
    doc.end();

    var iframe = document.getElementById("output-iframe");

    stream.on('finish', function () {
      iframe.src = stream.toBlobURL('application/pdf');
    });
  };
  reader.readAsBinaryString(f);

}

function formatExcelTime(time) {
  //let hours = Math.floor(time * 24);
  //let minutes = Math.floor(Math.abs(time * 24 * 60) % 60);
  let date = new Date(time * 864e5)
  return `${date.getHours()}:${date.getMinutes()}`;
}

function timeToNum(time) {
  let parts = time.split(":");
  if (parts.length < 2) {
    throw Error("Vstupní soubor obsahuje nesprávě zformátovaný čas");
  }
  return parts[0] / 24 + parts[1] / 60 / 24;
}

function timeStripSecs(time) {
  let parts = time.split(":");
  if (parts.length < 2) {
    throw Error("Vstupní soubor obsahuje nesprávě zformátovaný čas");
  }
  return `${parts[0]}:${parts[1]}`;
}

async function renderSheet(doc, data, className) {
  // Max min hours
  let timetableStart = 24;
  let timetableEnd = 0;
  data.forEach(subject => {
    let start = Math.floor(timeToNum(subject.Zacatek) * 24);
    let end = Math.ceil(timeToNum(subject.Konec) * 24);
    timetableStart = Math.min(timetableStart, start);
    timetableEnd = Math.max(timetableEnd, end);
  });
  console.log("timetable class", className);
  console.log("timetable start", timetableStart);
  console.log("timetable end", timetableEnd);

  TABLE_X1 = 20;
  TABLE_Y1 = 40;
  TABLE_X2 = 841.89 - TABLE_X1;
  TABLE_Y2 = 595.28 - TABLE_Y1;
  SUBJECTS_X = TABLE_X1 + 50;
  SUBJECTS_Y = TABLE_Y1 + 50;
  DAYS = ["Po", "Út", "St", "Čt", "Pá", "So", "Ne"]
  NDAYS = 5;

  doc.font("RobotoCondensed-Regular");

  doc.fontSize(30);
  doc
    .text(className, TABLE_X1, TABLE_X1, { ineBreak: false });

  doc.fontSize(12);
  // Render base layout
  for (let h = timetableStart; h <= timetableEnd; h++) {
    let cellWidth = ((TABLE_X2 - SUBJECTS_X) / (timetableEnd - timetableStart));
    let x = SUBJECTS_X + (h - timetableStart) * cellWidth;
    doc
      .save()
      .moveTo(x, SUBJECTS_Y)
      .lineTo(x, TABLE_Y2)
      .stroke(h == timetableStart || h == timetableEnd ? '#AAAAAA' : '#DDDDDD');
    let text = `${h}:00`;
    let textX = x;
    let textW = doc.widthOfString(text);
    doc
      .text(text, textX - textW / 2, SUBJECTS_Y - 10, { baseline: 'middle', lineBreak: false });
  }

  for (let i = 0; i < NDAYS + 1; i++) {
    let y = SUBJECTS_Y + i * (TABLE_Y2 - SUBJECTS_Y) / NDAYS;
    doc
      .save()
      .moveTo(TABLE_X1, y)
      .lineTo(TABLE_X2, y)
      .stroke('#AAAAAA');
    if (i < NDAYS) {
      let textY = y + (TABLE_Y2 - SUBJECTS_Y) / NDAYS / 2;
      doc
        .text(DAYS[i], TABLE_X1 + 5, textY, { baseline: 'middle', lineBreak: false });
    }
  }

  // Render subjects
  data.forEach(subject => {
    let start = (timeToNum(subject.Zacatek) * 24 - timetableStart) / (timetableEnd - timetableStart);
    let end = (timeToNum(subject.Konec) * 24 - timetableStart) / (timetableEnd - timetableStart);
    let day = DAYS.findIndex((element) => element == subject.Den);

    let height = (TABLE_Y2 - SUBJECTS_Y) / NDAYS;

    let x1 = SUBJECTS_X + start * (TABLE_X2 - SUBJECTS_X);
    let y1 = SUBJECTS_Y + day * height + height * 0.1;
    let w = (end - start) * (TABLE_X2 - SUBJECTS_X);
    let h = height * 0.8;
    //console.log(start, end, day, x1, y1, w, h);

    doc
      .save()
      .rect(x1, y1, w, h)
      .fill("#ffffff");
    doc
      .save()
      .rect(x1, y1, w, h)
      .stroke("#000000")
    doc.fill("#000000");
    doc
      .text(subject.Predmet, x1 + 5, y1 + 5, { lineBreak: false });
    doc
      .text(subject.Vyucujici, x1 + 5, y1 + 20, { lineBreak: false });
    doc
      .text(timeStripSecs(subject.Zacatek), x1 + 5, y1 + 35, { lineBreak: false });
    doc
      .text(timeStripSecs(subject.Konec), x1 + 5, y1 + 50, { lineBreak: false });
  });
}