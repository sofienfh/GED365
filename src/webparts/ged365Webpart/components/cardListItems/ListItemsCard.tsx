import React, { useState, useEffect } from 'react';
import { SPListItem, SPOperations } from "../../../Services/SPServices";
import { ActionButton, IContextualMenuItem, Modal, Button, TextField } from 'office-ui-fabric-react';
import styles from '../Ged365Webpart.module.scss';

interface IListItemsCardProps {
  context: any;
  liste_titre: string;
  items_number: number;
}

const ListItemsCard: React.FC<IListItemsCardProps> = ({ context, liste_titre, items_number }) => {
  const [listItems, setListItems] = useState<SPListItem[]>([]);
  const [showEditModal, setShowEditModal] = useState(false);
  const [Titre_list_item, setTitreListItem] = useState("");
  const [Item_Id, setItemId] = useState<string | null>(null);

  const _spOperations = new SPOperations();

  useEffect(() => {
    updateListItems();
  }, [liste_titre, items_number]);

  const updateListItems = () => {
    _spOperations.GetListItems(context, liste_titre)
      .then((results: SPListItem[]) => {
        setListItems(results);
        console.log("List items updated");
      })
      .catch(error => {
        console.error('Error updating list items:', error);
      });
  };

  const getAllKeysExcept = (obj: any, excludeKeys: string[]): string[] => {
    return Object.keys(obj).filter(key => !excludeKeys.includes(key));
  };

  const excludeKeys = ['ID', 'Id'];

  const renderTableRows = () => {
    if (listItems.length === 0) {
      return <p>No items found.</p>;
    }

    return (
      <div className={styles.cardsContainer}>
        {listItems.map((listItem, index) => {
          const keys = getAllKeysExcept(listItem, excludeKeys);
          const menuItems: IContextualMenuItem[] = [
            { key: 'edit', text: 'Edit', onClick: () => editItem(listItem.Id) },
            { key: 'delete', text: 'Delete', onClick: () => deleteItem(listItem.Id) },
          ];

          return (
            <div key={`card-${index}`} className={styles.card}>
              <div className={styles.card_header}>
                <h3>Liste {index + 1}</h3>
                <div className={styles.actionDropdown}>
                  <ActionButton
                    iconProps={{ iconName: 'MoreVertical' }}
                    title="Actions"
                    menuProps={{ items: menuItems }}
                  />
                </div>
              </div>
              <div className={styles.cardContent}>
                {keys.map((key, i) => {
                  const value = listItem[key];
                  if (value !== '' && value !== null && !key.startsWith('@odata') && !key.startsWith('OData') && !['FileSystemObjectType', 'ContentTypeId', 'Data__UIVersionString', 'GUID'].includes(key)) {
                    if (key.toLowerCase().includes('image')) {
                      let dataJson = JSON.parse(value);
                      return (
                        <img key={i} src={dataJson['serverUrl'] + dataJson['serverRelativeUrl']} alt={key} className={styles.cardContent} />
                      );
                    } else {
                      return (
                        <div key={i} className={styles.cardContent}>
                          <strong>{key}: </strong>{value}
                        </div>
                      );
                    }
                  }
                  return null;
                })}
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  const deleteItem = (itemId: string) => {
    console.log(itemId);
    if (itemId) {
      _spOperations.DeleteListItem(context, liste_titre, itemId)
        .then((result: string) => {
          console.log(result);
          updateListItems();
        })
        .catch(error => console.error('Error deleting list item:', error));
    }
  };

  const handleEditSubmit = () => {
    if (Item_Id) {
      _spOperations.UpdateListItem(context, liste_titre, Item_Id, Titre_list_item)
        .then((result: string) => {
          console.log(result);
          updateListItems();
        })
        .catch(error => console.error('Error updating list item:', error));
      toggleEditModal();
    }
  };

  const editItem = (itemId: string) => {
    console.log(`Edit item with ID: ${itemId}`);
    setItemId(itemId);
    toggleEditModal();
  };

  const toggleEditModal = () => {
    setShowEditModal(!showEditModal);
  };

  const handleTitleChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    setTitreListItem(event.target.value);
  };

  return (
    <div>
      <h2>Liste : {liste_titre}</h2>
      {renderTableRows()}
      <Modal
        isOpen={showEditModal}
        onDismiss={toggleEditModal}
        isBlocking={false}
        containerClassName={styles.modalContainer}
      >
        <div className={styles.modalHeader}>
          <span>Modal Header</span>
          <Button iconProps={{ iconName: 'Cancel' }} onClick={toggleEditModal} />
        </div>
        <div className={styles.modalBody}>
          <TextField label="Titre" onChange={handleTitleChange} value={Titre_list_item} />
        </div>
        <div className={styles.modalFooter}>
          <Button text="Submit" onClick={handleEditSubmit} />
        </div>
      </Modal>
    </div>
  );
};

export default ListItemsCard;
