import * as React from "react";
import { PrimaryButton, TextField } from "@fluentui/react";
import { Item } from "./ValidationItem";

import { setIconOptions } from '@fluentui/react/lib/Styling';

// Suppress icon warnings.
setIconOptions({
  disableWarnings: true
});

export interface Props {
  id?: any;
  item?: Item;
  update: (id: any, item: Item) => void;
  back: () => void;
}

export const ValidationConfig: React.FunctionComponent<Props> = ({ children, id, item, update, back }) => {
  const [config, setConfig] = React.useState<Item | null>(item);
  const [valid, setValid] = React.useState<boolean>(!!item);

  React.useEffect(() => {
    setValid(check());
  }, [config]);


  const onChange = React.useCallback(
    (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, id: string, newValue?: string) => {
      setConfig({ ...config, [id]: newValue });
    },
    [config],
  );

  const save = () => {
    update(id, { ...config, active: item?.active || true });
    back();
  }

  const check = () => {
    if (config) {
      const { title, primaryText, icon } = config;
      if (title && primaryText && icon) {
        return true;
      }
    }

    return false;
  }

  const render = () => {
    return (
      <main className="ms-welcome__main">
        {children}
        <h3 style={{ margin: "10px 0px" }} className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">{item?.title || "Nueva Validación"}</h3>
        <TextField value={config?.title || ""} onChange={(e, newValue) => onChange(e, "title", newValue)} required label="Nombre" />
        <TextField value={config?.icon || ""} onChange={(e, newValue) => onChange(e, "icon", newValue)} required label="Icono" iconProps={{ iconName: config?.icon }} />
        <TextField value={config?.primaryText || ""} onChange={(e, newValue) => onChange(e, "primaryText", newValue)} required label="Descripción" />
        <TextField value={config?.secondaryText || ""} onChange={(e, newValue) => onChange(e, "secondaryText", newValue)} label="Nota" />
        <PrimaryButton disabled={!valid} text="Guardar" onClick={save} style={{ width: "100%", margin: "10px 0px" }} iconProps={{ iconName: "Save" }} allowDisabledFocus />
      </main>
    );
  }

  return render();
};