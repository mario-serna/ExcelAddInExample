import * as React from "react";
import { DefaultButton, PrimaryButton } from "@fluentui/react";
import { Item, ValidationItem } from "./ValidationItem";
import { ValidationConfig } from "./ValidationConfig";

export interface Props {
  title: string;
}

const getItems = async (headers: { [key: string]: number }): Promise<Item[]> => {
  try {
    return await Excel.run(async (context) => {

      // Comprobar si existe la hoja, en caso contrario crearla.
      let validationWorksheet: Excel.Worksheet;
      try {
        validationWorksheet = context.workbook.worksheets.getItem("Validations");
        await context.sync();
      } catch (error) {
        return [];
      }

      let currentRange = validationWorksheet.getUsedRange();
      currentRange.load("values");
      let lastRow = currentRange.getLastRow();
      lastRow.load("rowIndex");
      await context.sync();
      let items = [];
      for (let i = 0; i < currentRange.values.length; i++) {
        if (i === 0) continue;
        let tempItem = {};
        const headerRow = currentRange.values[0];
        const row = currentRange.values[i];
        for (let j = 0; j < row.length; j++) {
          tempItem[headerRow[j]] = row[j];
        }
        items.push(tempItem);
      }
      console.log(items, headers);

      return items;
    });
  } catch (error) {
    console.error(error);
    return [];
  }
};

const saveItems = async (headers: { [key: string]: number }, items: Item[]) => {
  try {
    await Excel.run(async (context) => {

      // Comprobar si existe la hoja, en caso contrario crearla.
      let validationWorksheet: Excel.Worksheet;
      try {
        validationWorksheet = context.workbook.worksheets.getItem("Validations");
        await context.sync();
      } catch (error) {
        console.log(error)
        validationWorksheet = context.workbook.worksheets.add("Validations");

      }

      const headersNum = Object.keys(headers).length;
      const headersRange = validationWorksheet.getRangeByIndexes(0, 0, 1, headersNum);
      await context.sync();

      let valuesHeaders = [[]];
      for (const key in headers) {
        valuesHeaders[0][headers[key]] = key;
      }
      headersRange.values = valuesHeaders;

      if (items.length > 0) {
        for (let i = 0; i < items.length; i++) {
          let values = [[]];
          for (const key in items[i]) {
            values[0][headers[key]] = items[i][key];
          }
          const cell = validationWorksheet.getRangeByIndexes(i + 1, 0, 1, headersNum);
          cell.values = values;
        }
      } else {
        const delRange = validationWorksheet.getRangeByIndexes(items.length + 1, 0, 50, headersNum);
        delRange.delete("Up");
      }
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
};

const delItem = async (id: any) => {
  try {
    await Excel.run(async (context) => {

      // Comprobar si existe la hoja, en caso contrario crearla.
      let validationWorksheet: Excel.Worksheet;
      try {
        validationWorksheet = context.workbook.worksheets.getItem("Validations");
        await context.sync();
      } catch (error) {
        console.log(error)
        validationWorksheet = context.workbook.worksheets.add("Validations");

      }

      let currentRange = validationWorksheet.getUsedRange();
      const row = currentRange.getRow(id);
      row.delete("Up");

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
};

const HEADERS = {
  title: 0,
  icon: 1,
  primaryText: 2,
  secondaryText: 3,
  active: 4
};

export const ValidationList: React.FunctionComponent<Props> = ({ children, title }) => {
  const [items, setItems] = React.useState<Item[]>([]);
  const [selectedItem, setSelectedItem] = React.useState<number | null>(null);

  React.useEffect(() => {
    const loadItems = async () => {
      const tempItems = await getItems(HEADERS);
      setItems(tempItems);
    }

    loadItems();
  }, []);


  const updateItem = (id: any, data: Item) => {
    let tempItem = items;
    if (items[id]) {
      tempItem[id] = data;
      setItems(tempItem);
    } else {
      tempItem = [...items, data];
      setItems(tempItem);
    }

    saveItems(HEADERS, tempItem);
  }

  const deleteItem = (id: any) => {
    setItems(items.filter((_val, i) => i !== id));
    delItem(id + 1);
  }

  const render = () => {
    if (selectedItem === -1) {
      return <ValidationConfig update={updateItem} back={() => setSelectedItem(null)}>
        <DefaultButton text="Regresar" onClick={() => setSelectedItem(null)} style={{ marginBottom: 10 }} iconProps={{ iconName: "Back" }} allowDisabledFocus />
      </ValidationConfig>
    }

    if (selectedItem !== null) {
      return <ValidationConfig id={selectedItem} item={items[selectedItem]} update={updateItem} back={() => setSelectedItem(null)}>
        <DefaultButton text="Regresar" onClick={() => setSelectedItem(null)} style={{ marginBottom: 10 }} iconProps={{ iconName: "Back" }} allowDisabledFocus />
      </ValidationConfig>
    }

    const listItems = items.map((item, index) => (
      <ValidationItem
        key={index}
        item={item}
        index={index}
        select={() => setSelectedItem(index)}
        update={updateItem}
        deleteItem={deleteItem}
      />));
    return (
      <main className="ms-welcome__main">
        <h2 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">{title}</h2>
        <PrimaryButton text="Nueva Validación" onClick={() => setSelectedItem(-1)} style={{ width: "100%", marginBottom: 10 }} iconProps={{ iconName: "Add" }} allowDisabledFocus />
        {listItems}
        {children}
      </main>
    );
  }

  return render();
};