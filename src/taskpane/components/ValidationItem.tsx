import * as React from "react";
import { DocumentCard, DocumentCardActions, DocumentCardActivity, DocumentCardDetails, DocumentCardPreview, DocumentCardTitle, DocumentCardType, IButton, IButtonProps, Stack, Toggle } from "@fluentui/react";
import { getTheme } from '@fluentui/react/lib/Styling';
const theme = getTheme();
const { palette, fonts } = theme;

const onActionClick = (action: string, ev: React.SyntheticEvent<HTMLElement>): void => {
  console.log(`You clicked the ${action} action`);
  ev.stopPropagation();
  ev.preventDefault();
};


export interface Item {
  icon: string;
  title: string;
  primaryText: string;
  secondaryText?: string;
  active?: boolean;
}

interface Props {
  index: number;
  item: Item;
  select: (id: any) => void;
  update: (id: any, item: Item) => void;
  deleteItem: (id: any) => void;
}

export const ValidationItem: React.FunctionComponent<Props> = ({
  index, item, select, update, deleteItem
}) => {
  const [enabled, setEnabled] = React.useState(item.active);
  function _onChange(_ev: React.MouseEvent<HTMLElement>) {
    console.log('toggle is ' + (!enabled ? 'checked' : 'not checked'));
    setEnabled(!enabled);
    update(index, { ...item, active: !enabled });
  }

  return <DocumentCard
    styles={{ root: { maxWidth: "100%" } }}
    aria-label="Document Card with icon. View and share files. Created by Aaron Reid a few minutes ago"
    type={DocumentCardType.compact}
    onClick={() => {
      console.log(item);
      select(index);
    }}
  >
    <DocumentCardPreview {...{
      key: index,
      previewImages: [
        {
          previewIconProps: {
            iconName: item.icon,
            styles: { root: { fontSize: fonts.xxLarge.fontSize, color: palette.white } },
          },
          width: 64,
        },
      ],
      styles: { previewIcon: { backgroundColor: enabled ? palette.themePrimary : palette.themeLighter } },
    }} />
    <Stack style={{ flex: 1 }}>
      <Stack horizontal>
        <DocumentCardTitle styles={{ root: { flex: 1 } }} title={item.title} shouldTruncate />
        <DocumentCardActions styles={{ root: { padding: "4px 0px" } }} actions={[
          {
            menuProps: {
              items: [
                {
                  key: 'toggle',
                  iconProps: { iconName: enabled ? 'CheckboxComposite' : 'Checkbox' },
                  ariaLabel: 'share action',
                  checked: enabled,
                  text: enabled ? 'Activo' : 'Inactivo',
                  onClick: _onChange,
                  toggle: true,
                },
                {
                  key: 'calendarEvent',
                  text: 'Eliminar',
                  iconProps: { iconName: 'Delete' },
                  onClick: () => deleteItem(index),
                  ariaLabel: 'delete action',
                },
              ],
            }
          }
        ]} />
      </Stack>
      <DocumentCardActivity activity={item.secondaryText} people={[{ name: item.primaryText, profileImageSrc: '', initials: `${index}` }]} />
    </Stack>
  </DocumentCard>
}