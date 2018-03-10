var time_execution = new Date();

var formatDate = function (formatDate, formatString) {
    if(formatDate instanceof Date) {
        var months = new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec");
        var yyyy = formatDate.getFullYear();
        var yy = yyyy.toString().substring(2);
        var m = formatDate.getMonth() + 1;
        var mm = m < 10 ? "0" + m : m;
        var mmm = months[m];
        var d = formatDate.getDate();
        var dd = d < 10 ? "0" + d : d;

        var h = formatDate.getHours();
        var hh = h < 10 ? "0" + h : h;
        var n = formatDate.getMinutes();
        var nn = n < 10 ? "0" + n : n;
        var s = formatDate.getSeconds();
        var ss = s < 10 ? "0" + s : s;

        formatString = formatString.replace(/yyyy/i, yyyy);
        formatString = formatString.replace(/yy/i, yy);
        formatString = formatString.replace(/mmm/i, mmm);
        formatString = formatString.replace(/mm/i, mm);
        formatString = formatString.replace(/m/i, m);
        formatString = formatString.replace(/dd/i, dd);
        formatString = formatString.replace(/d/i, d);
        formatString = formatString.replace(/hh/i, hh);
        formatString = formatString.replace(/h/i, h);
        formatString = formatString.replace(/nn/i, nn);
        formatString = formatString.replace(/n/i, n);
        formatString = formatString.replace(/ss/i, ss);
        formatString = formatString.replace(/s/i, s);
        //Report.Debug("fdate=" + formatDate.toString() + " d=" + d.toString() + " h=" + h.toString() + " n=" + n.toString() + " s=" + s.toString());
        return formatString;
    } else {
        return "";
    }
}

var formatTime = function (sftime, formatString) {
    var sign = sftime < 0 ? "-" : "";
    var ftime = Math.abs(sftime);
    var d = Math.floor(ftime);
    var t = Math.floor(ftime * 24); //Всего часов
    var h = Math.floor((ftime - d) * 24);
    var n = Math.floor(((ftime - d) * 24 - h) * 60);
    var s = Number(((((ftime - d) * 24) - h) * 60 - n) * 60).toFixed();
    if (s == 60)
    {
        s = 0;
        n = n + 1;
        if (n == 60)
        {
            n = 0;
            h = h + 1;
            t = t + 1;
            if (h == 24)
            {
                h = 0;
                d = d + 1;
            }
        }
    }
    var dd = d < 10 ? "0" + d : d;
    var tt = t < 10 ? "0" + t : t;
    var hh = h < 10 ? "0" + h : h;
    var nn = n < 10 ? "0" + n : n;
    var ss = s < 10 ? "0" + s : s;
    //Report.Debug("ftime=" + ftime.toString() + " d=" + d.toString() + " h=" + h.toString() + " t=" + t.toString() + " n=" + n.toString() + " s=" + s.toString());

    formatString = formatString.replace(/dd/i, dd);
    formatString = formatString.replace(/d/i, d);
    formatString = formatString.replace(/tt/i, tt);
    formatString = formatString.replace(/t/i, t);
    formatString = formatString.replace(/hh/i, hh);
    formatString = formatString.replace(/h/i, h);
    formatString = formatString.replace(/nn/i, nn);
    formatString = formatString.replace(/n/i, n);
    formatString = formatString.replace(/ss/i, ss);
    formatString = formatString.replace(/s/i, s);

    return sign + formatString;
}

function SearchCoord(npos, obj) {
    var is_coord = false;
    var icoord;
    var v_lon=0;
    var v_lat=0;
    var result;
    var poscnt = obj.PositionCount;

    icoord = npos;
    while (!is_coord && (icoord > 0))
    {
        try
        {
            v_lon = obj.GetPosition(icoord).Lon;
            v_lat = obj.GetPosition(icoord).Lat;
        } catch(e) {}
        if ((v_lon == 0) || (v_lat == 0)) {icoord--} else {is_coord = true}
    }
    if (!is_coord) {icoord = npos + 1}
    while (!is_coord && (icoord < poscnt-1))
    {
        try
        {
            v_lon = obj.GetPosition(icoord).Lon;
            v_lat = obj.GetPosition(icoord).Lat;
        } catch(e) {}
        if ((v_lon == 0) || (v_lat == 0)) {icoord++} else {is_coord = true}
    }
    if (is_coord) {result = icoord} else {result = -1};
    return result;
}

Date.prototype.toOnlyDay = function () {
    return new Date(this.getFullYear(), this.getMonth(), this.getDate())
}

function addZone (ws, Zone, row, rate) {
    row++;
    var col_ = 0;
    ws.Cell(row, ++col_).Value = Zone.name;
    ws.Cell(row, ++col_).Value = '\'' + formatDate(new Date(Zone.enter), 'dd.mm.yyyy hh:nn:ss');
    ws.Cell(row, ++col_).Value = Zone.icoordBegin;
    ws.Cell(row, ++col_).Value = '\'' + formatDate(new Date(Zone.leave), 'dd.mm.yyyy hh:nn:ss');
    ws.Cell(row, ++col_).Value = Zone.icoordEnd;
    ws.Cell(row, ++col_).Value = '\'' + formatTime((Zone.leave - Zone.enter) / 24 / 3600000, 'tt:nn:ss');
    ws.Cell(row, ++col_).Value = Number(Zone.distEnd - Zone.distBegin).toFixed(2).replace('.',',');
    ws.Cell(row, ++col_).Value = Number((Zone.distEnd - Zone.distBegin) * rate).toFixed(2).replace('.',',');
    return row;

}
function zoneEnter() {
    var arr = {};
    arr.enter = v_time;
    arr.leave = v_time;
    arr.day = new Date(v_time).toOnlyDay();
    geozone = geozones.Find(pos.TaxCode);
    arr.name = geozone.Name;
    arr.distBegin = v_dist;
    arr.distEnd = v_dist;
    arr.icoordBegin = icoord;
    arr.icoordEnd = icoord;
    return arr;
}
function zoneLeave(){
    Zone.leave = v_time;
    Zone.distEnd = v_dist;
    Zone.icoordEnd = icoord;
    row = addZone (ws, Zone, row, rate);
    if(debug){
        var arr = {};
        arr.name = obj.Name;
        arr.zone = Zone;
        objs.push(arr);
        arr = {};
    }
    Zone = {};
    Zone.enter = 0;
}

var eng = Report.Engine;
var geozones = eng.Geofences;
var rate = Number(Report.GetParam("Parameters.Rate"));
var debug = false;
var objs = [];

var intper = 0;
var per;

var devCPU = 0x12;
var objcnt = Report.ObjectCount;
var obj;
var poscnt;
var pos;

var i, j, k, l;

var wb = new ActiveXObject("ClosedXML.Excel.XLWorkbook");
var svc = new ActiveXObject("ClosedXML.Excel.XLService");
var Color1 = svc.ColorFromArgb(0, 215, 228, 242);
var Color2 = svc.ColorFromArgb(0, 255, 182, 193);
var ColorFont = svc.ColorFromArgb(0, 54, 124, 209);
var ws = wb.Worksheets.Add('Нарушения');

ws.Cell(1, 4).Value = 'Отчет по системе «Платон»';
ws.Cell(1, 4).Style.Font.FontSize = 20; //размер шрифта
ws.Cell(1, 4).Style.Font.Bold = true; //жирный
ws.Cell(1, 4).Style.Font.FontColor = ColorFont;
ws.Cell(2, 4).Value = 'за период  с ' + formatDate(new Date(Report.Time1), 'dd.mm.yyyy hh:nn:ss') + ' по '
    + formatDate(new Date(Report.Time2), 'dd.mm.yyyy hh:nn:ss');
ws.Cell(2, 4).Style.Font.FontSize = 14; //размер шрифта
ws.Range_4(1, 4, 2, 4).Style.Alignment.Horizontal = 0;
var header_row = 3;
var col = 0;
ws.Cell(header_row, ++col).Value = 'Фед. трасса';
ws.Column(col).Width = 21;
ws.Cell(header_row, ++col).Value = 'Время входа';
ws.Column(col).Width = 10;
ws.Cell(header_row, ++col).Value = 'Координаты входа';
ws.Column(col).Width = 13;
ws.Cell(header_row, ++col).Value = 'Время выхода';
ws.Column(col).Width = 10;
ws.Cell(header_row, ++col).Value = 'Координаты выхода';
ws.Column(col).Width = 13;
ws.Cell(header_row, ++col).Value = 'Время нахождения';
ws.Column(col).Width = 12;
ws.Cell(header_row, ++col).Value = 'Пробег (км)';
ws.Column(col).Width = 7;
ws.Cell(header_row, ++col).Value = 'Стоимость (р)';
ws.Column(col).Width = 10;
function headerStyle(ws){
    ws.Range_4(header_row, 1, header_row, col).Style.Alignment.SetWrapText();
    ws.Range_4(header_row, 1, header_row, col).Style.Alignment.Horizontal = 0;
    ws.Range_4(header_row, 1, header_row, col).Style.Border.InsideBorder = 13; //совсем тонкий
    ws.Range_4(header_row, 1, header_row, col).Style.Border.OutsideBorder = 12;
}


var row_start = 3;
var row = row_start;
for (i = 0; i < objcnt; i++) {
    obj = Report.GetObject(i);
    poscnt = obj.PositionCount;
    var v_lon, v_lat;
    var Distance = 0;
    var v_time;
    var icoord;
    var taxCode;
    var Zone = {};
    Zone.enter = 0;

    var geozone;
    ws.Cell(++row, 4).Value = obj.Name + ' / ' + obj.AvtoNo;
    ws.Cell(row, 4).Style.Font.Bold = true; //жирный
    ws.Cell(row, 4).Style.Alignment.Horizontal = 0;
    ws.Range_4(row, 1, row, col).Style.Fill.BackgroundColor = Color1;
    ws.Range_4(row, 1, row, col).Style.Border.OutsideBorder = 13;
    var row_start_obj = row;

    for (j = 0; j < poscnt; j++) {
        pos = obj.GetPosition(j);
        if (pos.IsEvent(devCPU, 0, 0x3)) continue;  // evRestart on devCPU
        v_time = pos.Time;
        var v_state = pos.State;
        if (pos.Lon) v_lon = pos.Lon;
        if (pos.Lat) v_lat = pos.Lat;
        var v_dist = pos.Distance;

        if (j == 0) {
            taxCode = 0;
            time_prev = v_time;
        }

        if (taxCode != pos.TaxCode) {
            icoord = 'lon = ' + Number(v_lon).toFixed(5) + ' lat = ' + Number(v_lat).toFixed(5);
            if(taxCode != 0){ //завершение предыдущей и начало следующей / отсутсвие следующей
                //завершение предыдущей
                zoneLeave();
                //начало следующей, если есть
                if(pos.TaxCode != 0){
                    Zone = zoneEnter();
                }
            }else{
                //вход в зону из отсутсвия зоны
                Zone = {};
                Zone = zoneEnter();
            }
        }

        taxCode = pos.TaxCode;
        time_prev = v_time;

        // Show progress
        if (objcnt > 0 && poscnt > 0) {
            per = 100. * (1 / objcnt * (i + (j / poscnt)));
            if (Math.floor(per, 0) > Math.floor(intper, 0)) {
                Report.SetProgress(per);
                intper = per;
            }
        }
    }
    if(Zone.enter > 0){
        zoneLeave();
    }
    if(row_start_obj == row){
        row++;
        ws.Cell(row, 4).Value = 'Не посещено';
        ws.Range_4(row, 1, row, col).Style.Border.OutsideBorder = 13;
        ws.Cell(row, 4).Style.Alignment.Horizontal = 0;
        ws.Range_4(row, 1, row, col).Style.Fill.BackgroundColor = Color2;
    }else{
        ws.Range_4(row_start_obj + 1, 1, row, col).Style.Alignment.SetWrapText();
        ws.Range_4(row_start_obj + 1, 2, row, col).Style.Alignment.Horizontal = 0;
        ws.Range_4(row_start_obj + 1, 1, row, col).Style.Border.InsideBorder = 13; //совсем тонкий
        ws.Range_4(row_start_obj + 1, 1, row, col).Style.Border.OutsideBorder = 13;
    }

}

headerStyle(ws);

ws.PageSetup.PrintAreas.Add(1, 1, row, col);
ws.SheetView.FreezeRows(header_row);

ws.PageSetup.SetRowsToRepeatAtTop_2(header_row, header_row);
ws.PageSetup.PageOrientation = 1; //XLPageOrientation.Landscape
ws.PageSetup.PagesWide = 1;//сжимает по ширине область печати так, чтобы уместилась на печатной странице
ws.PageSetup.Header.Right.AddText_2(time_execution, 0);
ws.PageSetup.Header.Right.AddText_2("  Лист ", 0);
ws.PageSetup.Header.Right.AddText_3(0, 0);//XLHFPredefinedText.PageNumber, XLHFOccurrence.AllPages
ws.PageSetup.Header.Right.AddText_2(" / ", 0); //XLHFOccurrence.AllPages
ws.PageSetup.Header.Right.AddText_3(1, 0); //XLHFPredefinedText.NumberOfPages, XLHFOccurrence.AllPages
ws.PageSetup.Margins.Top = 0.5; //поле для печати сверху (отступ), где 1 = 2,54 см
ws.PageSetup.Margins.Bottom = 0.39; //поле печати снизу
ws.PageSetup.Margins.Left = 0.6; //поля в дюймах
ws.PageSetup.Margins.Right = 0.2;
ws.PageSetup.Margins.Header = 0.30; //поле верхнего колонтитула
ws.PageSetup.Margins.Footer = 0.28; //поле нижнего колонтитула
//ws.Cell(_row, _column).Style.Font.FontSize = 12; //размер шрифта

wb.SaveAs(Report.GetFileName());
/*
if (maps.length > 0) {//прибить файлы вставленных карт
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    for (i = 0; i < maps.length; i++) {
        if (fso.FileExists(maps[i])) fso.DeleteFile(maps[i])
    }
}
*/
