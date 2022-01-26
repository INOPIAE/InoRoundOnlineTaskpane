/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("iunround").onclick = iunround;
    document.getElementById("iround").onclick = iround;
  }
});

export async function iround() {
  try {
    await Excel.run(async (context) => {
      let myRng = context.workbook.getSelectedRange();
      myRng.load(["values", "text", "formulas", "formulasR1C1"]);
/*      var checkBoxNum = document.getElementById("inochknum");
      console.log(checkBoxNum);
      console.log(checkBoxNum.checked);*/

      // define digits
      var digits = document.getElementById("inoDigits");
      var digit = 2;
      //ToDo check digit isNumeric
      if (digits.value.indexOf(".") < 0) {
        digit = -digits.value.length + 1;
      } else {
        digit = digits.value.length - 2;
      }

      await context.sync();
      let myFormulas = myRng.formulasR1C1;

      let myOuterArray = [];

      // loop through each row
      myFormulas.forEach(function (row) {
        // define the inner array
        let myInnerArray = [];

        // then loop through each column
        row.forEach(function (col) {
          //console.log(col);
          let colString = col.toString();
          let test = colString.substring(0, 1);

          if (test == "=") {
            let newFormula = colString.replace("=", "=ROUND(") + ", " + digit + ")";
            console.log(newFormula);
            if (colString.length >= 6) {
              let test = colString.substring(0, 6);
              if (test == "=ROUND") {
                myInnerArray.push(col);
              } else {
                myInnerArray.push(newFormula);
              }
            } else {
              myInnerArray.push(newFormula);
            }
/*          } else if (checkBoxNum.checked == true) {
            if (isNaN(col) == true) {
              let newFormula = "=ROUND(" + colString + ", 2)";
              console.log(newFormula);
              //myInnerArray.push(newFormula);
            }*/
          } else {
            myInnerArray.push(col);
          }
        });

        // append the inner array to the outer array
        myOuterArray.push(myInnerArray);
      });

      //replace orginal value with neu values
      myRng.formulasR1C1 = myOuterArray;
    });
  } catch (error) {
    console.error(error);
  }
}

export async function iunround() {
  try {
    await Excel.run(async (context) => {
      let myRng = context.workbook.getSelectedRange();
      myRng.load(["values", "text", "formulas", "formulasR1C1"]);

      await context.sync();
      let myFormulas = myRng.formulasR1C1;

      let myOuterArray = [];

      // loop through each row
      myFormulas.forEach(function (row) {
        // define the inner array
        let myInnerArray = [];

        // then loop through each column
        row.forEach(function (col) {
          let colString = col.toString();
          let test = colString.substring(0, 6);

          if (test == "=ROUND") {
            let newFormula = colString.replace("=ROUNDDOWN(", "=");
            newFormula = newFormula.replace("=ROUNDUP(", "=");
            newFormula = newFormula.replace("=ROUND(", "=");
            newFormula = newFormula.substring(0, newFormula.lastIndexOf(","));
            myInnerArray.push(newFormula);
          } else {
            myInnerArray.push(col);
          }
        });

        // append the inner array to the outer array
        myOuterArray.push(myInnerArray);
      });

      //replace orginal value with neu values
      myRng.formulasR1C1 = myOuterArray;
    });
  } catch (error) {
    console.error(error);
  }
}
