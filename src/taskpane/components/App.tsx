import * as React from "react";
// import Header from "./Header";
// import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { ValidationList } from "./ValidationList";
import { Item } from "./ValidationItem";

/* global console, Excel, require  */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const App: React.FC<AppProps> = ({ title, isOfficeInitialized }) => {

  const click = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();
        // Read the range address
        range.load(["address", "values", "text"]);

        // Update the fill color
        range.format.fill.color = "yellow";
        range.format.font.bold = true;
        await context.sync();

        range.values = [["Texto"]];
        console.log(`The range address was ${range.address}.`);
        console.log(`The range address was ${range.values}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };

  const createParamsWorksheet = async () => {
    try {
      await Excel.run(async (context) => {

        // Comprobar si existe la hoja, en caso contrario crearla.
        let paramsWorksheet: Excel.Worksheet;
        try {
          paramsWorksheet = context.workbook.worksheets.getItem("Params");
          await context.sync();
        } catch (error) {
          console.log(error)
          paramsWorksheet = context.workbook.worksheets.add("Params");

        }

        // Insertar las etiquetas de las columnas
        const h = paramsWorksheet.getRange("A1:C1");
        h.values = [["X", "Y", "Z"]];
        h.format.fill.color = "#000000";
        h.format.font.color = "#ffffff";
        h.format.horizontalAlignment = "Center";

        // Insertar los valores
        const r = paramsWorksheet.getRange("A2:C3");
        r.load(["values", "text"]);
        await context.sync();

        let values = [];
        for (let i = 0; i < r.values.length; i++) {
          const row = r.values[i];
          values[i] = [];
          for (let j = 0; j < row.length; j++) {
            // row[j] = i + j; // No es posible modificar el valor
            values[i][j] = (i + j) * 10;
          }
        }

        console.log(r.values);

        r.values = values;
        r.format.fill.color = "#4472c2";
        r.format.font.color = "#ffffff";

        // Insertar formulas
        const f = paramsWorksheet.getRange("A4:C4");
        f.formulas = [["=SUM(A2:A3)", "=SUM(B2:B3)", "=SUM(C2:C3)"]];
        f.format.fill.color = "yellow";
        f.format.font.bold = true;

        // f.load("formulas"); // Cargar solo para obtener los valores
        // r.values = [[1, 2, 3], [4, 5, 6]];
        // r.values[0] = [11, 22, 33];
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  };

  const render = () => {

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        {/* <Header logo={require("./../../../assets/logo-filled.png")} title={title} message="Welcome" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={click}>
            Run
          </DefaultButton>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={createParamsWorksheet}>
            Add Params
          </DefaultButton>
        </HeroList> */}

        <ValidationList title="Mis Validaciones" />
      </div>
    );
  }

  return render();
}

export default App;