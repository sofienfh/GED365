import { useState, useEffect } from 'react';
import { SPListColumn, SPListItem, SPOperations } from "../../../Services/SPServices";
import { CommandBarButton } from 'office-ui-fabric-react';
import { Modal, TextField } from 'office-ui-fabric-react';
import React from 'react';
import styles from '../Ged365Webpart.module.scss';
import 'bootstrap/dist/css/bootstrap.min.css';

interface IListItemsRendererProps {
    context: any;
    liste_titre: string;
    items_number: number;
}

const ListItemsRenderer: React.FC<IListItemsRendererProps> = ({ context, liste_titre, items_number }) => {
    const [listItems, setListItems] = useState<SPListItem[]>([]);
    const [colonnes, setColonnes] = useState<SPListColumn[]>([]);
    const [showEditModal, setShowEditModal] = useState(false);
    const [Titre_list_item, setTitreListItem] = useState("");
    const [Item_Id, setItemId] = useState<string | null>(null);

    const _spOperations = new SPOperations();

    useEffect(() => {
        updateListItems();
    }, [liste_titre, items_number]);

    const shouldRenderItem = (key: string, value: any) => {
        return (
            !key.startsWith('@odata') &&
            !key.startsWith('OData') &&
            key !== 'FileSystemObjectType' &&
            !['ContentTypeId', 'Data__UIVersionString', 'GUID'].includes(key) &&
            value !== "" &&
            value !== null &&
            value !== false &&
            key !== undefined
        );
    };

    const updateListItems = () => {
        _spOperations.GetListItems(context, liste_titre)
            .then((results: SPListItem[]) => {
                setListItems(results);
                console.log("List items updated");
            })
            .catch(error => {
                console.error('Error updating list items:', error);
            });

        _spOperations.GetListColumns(context, liste_titre)
            .then((results: SPListColumn[]) => {
                setColonnes(results);
                console.log("List columns updated");
            })
            .catch(error => {
                console.error('Error getting list columns:', error);
            });
    };

    const renderTableHeader = () => {
        if (colonnes.length === 0) return null;

        return (
            <tr>
                {colonnes.map(column => {
                    if (shouldRenderItem(column.internalName, column.title)) {
                        return <th key={column.internalName}>{column.title}</th>;
                    }
                    return null;
                })}
                <th>Action</th>
            </tr>
        );
    };

    const renderTableRows = () => {
        if (listItems.length === 0) {
            return <tr><td colSpan={colonnes.length + 1}>No items found</td></tr>;
        }

        return listItems.map((listItem, index) => {
            return (
                <tr key={`list-${index}`}>
                    {colonnes.map(column => {
                        const value = listItem[column.internalName];
                        return <td key={column.internalName}>{value ? value : "--"}</td>;
                    })}
                    <td>
                        <CommandBarButton
                            menuIconProps={{ iconName: 'More' }}
                            menuProps={{
                                items: [
                                    {
                                        key: 'edit',
                                        text: 'Edit',
                                        iconProps: { iconName: 'Edit' },
                                        onClick: () => editItem(listItem.Id)
                                    },
                                    {
                                        key: 'delete',
                                        text: 'Delete',
                                        iconProps: { iconName: 'Delete' },
                                        onClick: () => deleteItem(listItem.Id)
                                    }
                                ]
                            }}
                        />
                    </td>
                </tr>
            );
        });
    };

    const deleteItem = (itemId: string) => {
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
            <div className="overflow-x-scroll">
                <table className="table">
                    <thead>{renderTableHeader()}</thead>
                    <tbody>{renderTableRows()}</tbody>
                </table>
            </div>
            <Modal
                isOpen={showEditModal}
                onDismiss={toggleEditModal}
                isBlocking={false}
                containerClassName={styles.modalContainer}
            >
                <div className={styles.modalHeader}>
                    <span>Modal Header</span>
                    <CommandBarButton
                        iconProps={{ iconName: 'Cancel' }}
                        onClick={toggleEditModal}
                    />
                </div>
                <div className={styles.modalBody}>
                    <TextField label="Titre" onChange={handleTitleChange} value={Titre_list_item} />
                </div>
                <div className={styles.modalFooter}>
                    <CommandBarButton
                        text="Submit"
                        onClick={handleEditSubmit}
                    />
                </div>
            </Modal>
        </div>
    );
};

export default ListItemsRenderer;
