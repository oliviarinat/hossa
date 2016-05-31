$(function() {
  var handleDrop = function (e) {
    e.stopPropagation();
    e.preventDefault();
    //http://stackoverflow.com/questions/14788862/drag-drop-doenst-not-work-dropeffect-of-undefined
    var files = e.originalEvent.dataTransfer.files;
    var i,f;
    for (i = 0, f = files[i]; i != files.length; ++i) {
      var reader = new FileReader();
      var name = f.name;
      var LINE_SEP = "\n";
      reader.onload = function(e) {
        var data = e.target.result, vcard = "";

        /* if binary string, read with type 'binary' */
        var workbook = XLSX.read(data, {type: 'binary'});
        var first_sheet_name = workbook.SheetNames[0];
        /* Get worksheet */
        var worksheet = workbook.Sheets[first_sheet_name];

        var row = 2, names, values, value, i = 0;

        for (row = 2; row <= 23; row++) {
          //Name
          names = worksheet["A" + row].v.split("\n");
          for (i = 0; i < names.length; i++) {
            vcard += "BEGIN:VCARD" + LINE_SEP + "VERSION:4.0" + LINE_SEP;
            value = names[i];
            vcard += "N:" + value.replace(",", ";") + LINE_SEP;

            //Kinder
            value = worksheet["B" + row].v;
            vcard += "NOTE:" + value + " (" + worksheet["C" + row].v + ")";
            value = worksheet["D" + row] ? worksheet["D" + row].v : "";
            if (value) {
              vcard += ", " + value + " (" + worksheet["E" + row].v + ")";;
            }
            vcard += LINE_SEP;

            //Gruppen
            value = worksheet["C" + row].v;
            vcard += "CATEGORIES:" + value;
            value = worksheet["E" + row] ? worksheet["E" + row].v : "";
            if (value) {
              vcard += ", " + value;
            }
            vcard += LINE_SEP;

            //ADR;TYPE=home:;;Heidestrasse 17;Koeln;;51147;Germany
            vcard += "ADR;TYPE=private:;;";
            //Straße
            vcard += worksheet["F" + row].v + ";";
            //PLZ-Stadt
            values = worksheet["G" + row].v.split(" ");
            vcard += values[1] + ";;" + values[0] + ";;" + LINE_SEP;

            //E-Mail
            values = worksheet["H" + row].v.split("\n");
            if (i == 0) {
              vcard += "EMAIL:" + values[0] + LINE_SEP;
            } else if (values.length > 1) {
              vcard += "EMAIL:" + values[1] + LINE_SEP;
            }

            //Festnetz
            //TEL;TYPE=cell:(0170) 1234567
            if (worksheet["I" + row]) {
              vcard += "TEL;TYPE=home:" + worksheet["I" + row].v + LINE_SEP;
            }

            //Handy
            if (worksheet["J" + row]) {
              values = worksheet["J" + row].v.split("\n");
              if (i == 0) {
                vcard += "TEL;TYPE=cell:" + values[0] + LINE_SEP;
              } else if (values.length > 1) {
                vcard += "TEL;TYPE=cell:" + values[1] + LINE_SEP;
              }
            }

            vcard += "END:VCARD" + LINE_SEP;
          }
        }



        window.open("data:text/vcard;charset=utf-8," + encodeURI(vcard));
      };
      reader.readAsBinaryString(f);
    }
  };

  $("#dropArea").on("drop", handleDrop);
  $("#dropArea").on("dragover", function (e) {
    e.preventDefault();
  });

});
