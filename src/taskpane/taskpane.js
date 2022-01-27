/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("iunround").onclick = iunround;
    document.getElementById("iround").onclick = iround;
    document.getElementById("iroundup").onclick = iroundup;
    document.getElementById("irounddown").onclick = irounddown;
  }
});

export async function iround() {
  irounding("");
}

export async function iroundup() {
  irounding("up");
}

export async function irounddown() {
  irounding("down");
}

export async function irounding(roundType) {
  try {
    await Excel.run(async (context) => {
      let myRng = context.workbook.getSelectedRange();
      myRng.load(["values", "text", "formulas", "formulasR1C1"]);
      var checkBoxNum = document.getElementById("inochknum");

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
      let myValues = myRng.values;
      let myText = myRng.text;

      let myOuterArray = [];

      let rounding = "=ROUND(";
      switch (roundType) {
        case "up":
          rounding = "=ROUNDUP(";
          break;
        case "down":
          rounding = "=ROUNDDOWN(";
          break;
        default:
          rounding = "=ROUND(";
      }

      var cRow = 0;
      var cCol = 0;
      // loop through each row
      myFormulas.forEach(function (row) {
        // define the inner array
        let myInnerArray = [];

        // then loop through each column
        row.forEach(function (col) {
          let colString = col.toString();
          let test = colString.substring(0, 1);
          if (test == "=") {
            let newFormula = colString.replace("=", rounding) + ", " + digit + ")";
            if (colString.length >= 6) {
              let test = colString.substring(0, 6);
              if (test == "=ROUND") {
                let newFormula1 = RemoveRound(colString);
                newFormula = newFormula1.replace("=", rounding) + ", " + digit + ")";
                myInnerArray.push(newFormula);
              } else {
                myInnerArray.push(newFormula);
              }
            } else {
              myInnerArray.push(newFormula);
            }
          } else if (checkBoxNum.checked == true) {
            if (isNumeric(myValues[cRow][cCol]) == true && isDate(myText[cRow][cCol]) == false) {
              let newFormula = rounding + colString + ", " + digit + ")";
              console.log(newFormula);
              myInnerArray.push(newFormula);
            } else {
              myInnerArray.push(col);
            }
          } else {
            myInnerArray.push(col);
          }
          ++cCol;
        });

        // append the inner array to the outer array
        myOuterArray.push(myInnerArray);
        ++cRow;
        cCol = 0;
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
            let newFormula = RemoveRound(colString);
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

function RemoveRound(OldFormula) {
  var newFormula = OldFormula.replace("=ROUNDDOWN(", "=");
  newFormula = newFormula.replace("=ROUNDUP(", "=");
  newFormula = newFormula.replace("=ROUND(", "=");
  newFormula = newFormula.substring(0, newFormula.lastIndexOf(","));
  return newFormula;
}

function isNumeric(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}

function isDate(value) {
  switch (typeof value) {
    case "number":
      return true;
    case "string":
      return !isNaN(Date.parse(value));
    case "object":
      if (value instanceof Date) {
        return !isNaN(value.getTime());
      }
      return false;
    default:
      return false;
  }
}
