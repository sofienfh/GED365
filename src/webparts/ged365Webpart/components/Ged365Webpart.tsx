import * as React from 'react';
import {IDropdownOption, IChoiceGroupOption } from 'office-ui-fabric-react';
import styles from './Ged365Webpart.module.scss';
import { IGed365WebpartProps } from './IGed365WebpartProps';
import { IGed365WebpartState } from './IGed365WebpartState';
import { SPOperations, SPListColumn } from '../../Services/SPServices';
import ButtonGrid from './ButtonGrid';
import CreateDocumentModal from './CreateDocumentModal';
import UploadDocumentModal from './UploadDocumentModal';
import AddMetadataModal from './AddMetadataModal';
import CreateFolderModal from './CreateFolderModal';

const metadataTypeOptions: IDropdownOption[] = [
  { key: 'Text', text: 'Ligne de texte' },
  { key: 'Choice', text: 'Choix' },
  { key: 'Number', text: 'Nombre' },
  { key: 'Boolean', text: 'Oui/Non' },
  { key: 'Image', text: 'Image' },
];

const choiceGroupOptions: IChoiceGroupOption[] = [
  { key: 'Yes', text: 'Oui' },
  { key: 'No', text: 'Non' },
];

export default class Ged365Webpart extends React.Component<IGed365WebpartProps, IGed365WebpartState> {
  public _spOperations: SPOperations;

  constructor(props: IGed365WebpartProps) {
    super(props);
    this._spOperations = new SPOperations();
    this.state = {
      listColumns: [],
      listTiltes: [],
      listItems: [],
      status: '',
      Titre_list_item: '',
      showModal: false,
      listItemId: '',
      selectedDocumentType: 'txt',
      metadata: {},
      uploadFile: null,
      isUploadMode: false,
      showCreateModal: false,
      showUploadModal: false,
      showAddMetadataModal: false,
      newMetadataField: '',
      newMetadataDescription: '',
      newMetadataType: 'Text',
      choices: [''],
      showCreateFolderModal: false, // New state for folder creation modal
      newFolderName: '', // New state for storing the folder name
    };
  }

  public openCreateModal = () => {
    this.setState({ showCreateModal: true }, () => {
      if (this.props.list_title) {
        this._spOperations.GetListColumns(this.props.context, this.props.list_title)
          .then((results: SPListColumn[]) => {
            const metadata: { [key: string]: any } = {};
            results.forEach(column => {
              metadata[column.internalName] = '';
            });
            this.setState({ listColumns: results, metadata });
          })
          .catch(error => {
            console.error('Error getting list columns:', error);
          });
      }
    });
  };

  public openUploadModal = () => {
    this.setState({ showUploadModal: true }, () => {
      if (this.props.list_title) {
        this._spOperations.GetListColumns(this.props.context, this.props.list_title)
          .then((results: SPListColumn[]) => {
            const metadata: { [key: string]: any } = {};
            results.forEach(column => {
              metadata[column.internalName] = '';
            });
            this.setState({
              listColumns: results.filter(column => column.internalName !== 'Nom'), // Exclure le champ "Nom"
              metadata,
            });
          })
          .catch(error => {
            console.error('Error getting list columns:', error);
          });
      }
    });
  };

  public openAddMetadataModal = () => {
    this.setState({ showAddMetadataModal: true });
  };

  public openCreateFolderModal = () => {
    this.setState({ showCreateFolderModal: true }, () => {
      if (this.props.list_title) {
        this._spOperations.GetListColumns(this.props.context, this.props.list_title)
          .then((results: SPListColumn[]) => {
            const metadata: { [key: string]: any } = {};
            results.forEach(column => {
              if (column.internalName !== 'Nom_du_dossier') { // Exclure spécifiquement le champ 'Nom du dossier'
                metadata[column.internalName] = '';
              }
            });
            this.setState({ listColumns: results, metadata });
          })
          .catch(error => {
            console.error('Error getting list columns:', error);
          });
      }
    });
  };

  private handleNewMetadataFieldChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    this.setState({ newMetadataField: event.target.value });
  };

  private handleNewMetadataDescriptionChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    this.setState({ newMetadataDescription: event.target.value });
  };

  private handleNewMetadataTypeChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (option) {
      this.setState({ newMetadataType: option.key as string, choices: [''] });
    }
  };

  private handleInputChange = (internalName: string) => (event: React.ChangeEvent<HTMLInputElement>) => {
    const value = event.target.type === 'checkbox' ? event.target.checked : event.target.value;
    this.setState(prevState => ({
      metadata: {
        ...prevState.metadata,
        [internalName]: value,
      },
    }));
  };

  private handleSelectChange = (internalName: string) => (event: React.ChangeEvent<HTMLSelectElement>) => {
    const value = event.target.value;
    this.setState(prevState => ({
      metadata: {
        ...prevState.metadata,
        [internalName]: value,
      },
    }));
  };

  private handleTextareaChange = (internalName: string) => (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    const value = event.target.value;
    this.setState(prevState => ({
      metadata: {
        ...prevState.metadata,
        [internalName]: value,
      },
    }));
  };

  private handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files && event.target.files[0];
    if (file) {
      this.setState({ uploadFile: file });
    }
  };

  private handleChoiceChange = (index: number) => (event: React.ChangeEvent<HTMLInputElement>) => {
    const newChoices = [...this.state.choices];
    newChoices[index] = event.target.value;
    this.setState({ choices: newChoices });
  };

  private addChoiceField = () => {
    this.setState(prevState => ({ choices: [...prevState.choices, ''] }));
  };

  private addMetadataField = () => {
    const newFieldName = this.state.newMetadataField;
    const newFieldType = this.state.newMetadataType;
    const choices = this.state.choices;

    if (this.props.list_title) {
      this._spOperations.AddMetadataField(this.props.context, this.props.list_title, newFieldName, newFieldType, choices)
        .then((result: string) => {
          console.log(result);
          this._spOperations.GetListColumns(this.props.context, this.props.list_title)
            .then((results: SPListColumn[]) => {
              const metadata: { [key: string]: any } = {};
              results.forEach(column => {
                metadata[column.internalName] = '';
              });
              this.setState({
                listColumns: results,
                metadata,
                newMetadataField: '',
                newMetadataDescription: '',
                newMetadataType: 'Text',
                showAddMetadataModal: false,
                choices: [''],
              });
            })
            .catch(error => {
              console.error('Error getting list columns:', error);
            });
        })
        .catch(error => {
          console.error('Error adding metadata field:', error);
        });
    }
  };

  private handleDocumentTypeChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (option) {
      this.setState({ selectedDocumentType: option.key as string });
    }
  };

  private handleCreateSubmit = () => {
    const fileType = this.state.selectedDocumentType;
    const metadata = { ...this.state.metadata };
    const fileName = `${metadata['Nom']}.${fileType}`; // Utiliser la métadonnée "Nom"

    this._spOperations.CreateFile(this.props.context, this.props.list_title, fileName, fileType, metadata)
      .then((result: string) => {
        this.setState({ status: result });
      })
      .catch(error => {
        console.error('Error creating file:', error);
      });

    this.setState({ showCreateModal: false, selectedDocumentType: 'txt' });
  };

  private handleUploadSubmit = () => {
    if (this.state.uploadFile) {
      const metadata = { ...this.state.metadata };
      const fileName = this.state.uploadFile.name;
      metadata['Nom'] = fileName; // Utiliser le nom du fichier téléchargé comme métadonnée "Nom"

      this._spOperations.UploadFile(this.props.context, this.props.list_title, this.state.uploadFile, metadata)
        .then((result: string) => {
          this.setState({ status: result });
        })
        .catch(error => {
          console.error('Error uploading file:', error);
        });

      this.setState({ showUploadModal: false, uploadFile: null });
    }
  };

  private handleCreateFolderSubmit = () => {
    const { newFolderName } = this.state;
    const metadata = { ...this.state.metadata };
    if (newFolderName) {
      this._spOperations.CreateFolder(this.props.context, this.props.list_title, newFolderName)
        .then(result => {
          // Update folder metadata
          this._spOperations.GetFolderItem(this.props.context, this.props.list_title, newFolderName)
            .then(folderItemId => {
              if (folderItemId) {
                this._spOperations.UpdateListItem(this.props.context, this.props.list_title, folderItemId, newFolderName, metadata)
                  .then(updateResult => {
                    alert(updateResult); // Success message
                    this.setState({ showCreateFolderModal: false, newFolderName: '', metadata: {} }); // Reset and close modal
                  })
                  .catch(error => {
                    console.error('Error updating folder metadata:', error);
                  });
              }
            })
            .catch(error => {
              console.error('Error getting folder item ID:', error);
            });
        })
        .catch(error => {
          console.error('Error creating folder:', error);
        });
    }
  };

  private getButtonClass = () => {
    switch (this.props.buttonType) {
      case 'rounded':
        return styles.buttonRounded;
      case 'semi-rounded':
        return styles.buttonSemiRounded;
      case 'strict':
        return styles.buttonStrict;
      default:
        return '';
    }
  };

  public render(): React.ReactElement<IGed365WebpartProps> {
    const { hasTeamsContext, backgroundColor, textColor } = this.props;

    const buttonStyle = {
      backgroundColor: backgroundColor,
      color: textColor, // Set the text color for the buttons
    };

    const documentTypeOptions: IDropdownOption[] = [
      { key: 'docx', text: 'Word Document (.docx)' },
      { key: 'txt', text: 'Text Document (.txt)' },
      { key: 'pptx', text: 'PowerPoint Presentation (.pptx)' },
      { key: 'xlsx', text: 'Excel Spreadsheet (.xlsx)' },
    ];

    const columnFields = this.state.listColumns
      .filter(column => !(this.state.isUploadMode && column.internalName === 'FileLeafRef'))
      .map(column => {
        let inputType: string | undefined;

        switch (column.type) {
          case 'Text':
            inputType = 'text';
            break;
          case 'Note':
            inputType = 'textarea';
            break;
          case 'Number':
            inputType = 'number';
            break;
          case 'DateTime':
            inputType = 'date';
            break;
          case 'Boolean':
            inputType = 'checkbox';
            break;
          case 'Choice':
            inputType = 'select';
            break;
          case 'URL':
            if (column.displayFormat === 1) {
              inputType = 'file';
            }
            break;
          default:
            inputType = 'text';
        }

        if (inputType === 'select') {
          return (
            <div key={column.internalName} className="mb-3">
              {!column.readOnly && (
                <>
                  <label htmlFor={column.internalName}>{column.title}</label>
                  <div className={styles['field-wrapper']}>
                    <div className={styles['field-group']}>
                      <select
                        name={column.internalName}
                        className={styles['select-field']}
                        id={column.internalName}
                        onChange={this.handleSelectChange(column.internalName)}
                      >
                        {column.choices &&
                          column.choices.map(choice => (
                            <option key={choice} value={choice}>
                              {choice}
                            </option>
                          ))}
                      </select>
                    </div>
                  </div>
                </>
              )}
            </div>
          );
        } else if (inputType === 'file') {
          return (
            <div key={column.internalName} className="mb-3">
              {!column.readOnly && (
                <>
                  <label htmlFor={column.internalName}>{column.title}</label>
                  <div className={styles['field-wrapper']}>
                    <div className={styles['field-group']}>
                      <input
                        type="file"
                        id={column.internalName}
                        accept="image/*"
                        className={styles['text-field']}
                        onChange={this.handleFileChange}
                      />
                    </div>
                  </div>
                </>
              )}
            </div>
          );
        } else if (inputType === 'checkbox') {
          return (
            <div key={column.internalName} className="mb-3">
              {!column.readOnly && (
                <div className={styles['field-wrapper']}>
                  <input
                    type="checkbox"
                    id={column.internalName}
                    className={styles['checkbox-field']}
                    onChange={this.handleInputChange(column.internalName)}
                    checked={!!this.state.metadata[column.internalName]}
                  />
                  <label htmlFor={column.internalName}>{column.title}</label>
                </div>
              )}
            </div>
          );
        } else if (inputType === 'textarea') {
          return (
            <div key={column.internalName} className="mb-3">
              {!column.readOnly && (
                <>
                  <label htmlFor={column.internalName}>{column.title}</label>
                  <div className={styles['field-wrapper']}>
                    <div className={styles['field-group']}>
                      <textarea
                        id={column.internalName}
                        className={styles['text-field']}
                        onChange={this.handleTextareaChange(column.internalName)}
                        value={this.state.metadata[column.internalName] || ''}
                      />
                    </div>
                  </div>
                </>
              )}
            </div>
          );
        } else {
          return (
            <div key={column.internalName} className="mb-3">
              {!column.readOnly && (
                <>
                  <label htmlFor={column.internalName}>{column.title}</label>
                  <div className={styles['field-wrapper']}>
                    <div className={styles['field-group']}>
                      <input
                        type={inputType}
                        id={column.internalName}
                        className={styles['text-field']}
                        onChange={this.handleInputChange(column.internalName)}
                        value={this.state.metadata[column.internalName] || ''}
                      />
                    </div>
                  </div>
                </>
              )}
            </div>
          );
        }
      });

    return (
      <section className={`${styles.ged365Webpart} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <ButtonGrid
            openCreateModal={this.openCreateModal}
            openUploadModal={this.openUploadModal}
            openAddMetadataModal={this.openAddMetadataModal}
            openCreateFolderModal={this.openCreateFolderModal}
            listTitle={this.props.list_title}
            buttonStyle={buttonStyle}
            getButtonClass={this.getButtonClass}
          />

          <CreateDocumentModal
            isOpen={this.state.showCreateModal}
            onDismiss={() => this.setState({ showCreateModal: false })}
            handleCreateSubmit={this.handleCreateSubmit}
            handleDocumentTypeChange={this.handleDocumentTypeChange}
            columnFields={columnFields}
            documentTypeOptions={documentTypeOptions}
          />

          <UploadDocumentModal
            isOpen={this.state.showUploadModal}
            onDismiss={() => this.setState({ showUploadModal: false })}
            handleUploadSubmit={this.handleUploadSubmit}
            handleFileChange={this.handleFileChange}
            columnFields={columnFields}
          />

          <AddMetadataModal
            isOpen={this.state.showAddMetadataModal}
            onDismiss={() => this.setState({ showAddMetadataModal: false })}
            handleNewMetadataFieldChange={this.handleNewMetadataFieldChange}
            handleNewMetadataDescriptionChange={this.handleNewMetadataDescriptionChange}
            handleNewMetadataTypeChange={this.handleNewMetadataTypeChange}
            handleChoiceChange={this.handleChoiceChange}
            addChoiceField={this.addChoiceField}
            addMetadataField={this.addMetadataField}
            newMetadataField={this.state.newMetadataField}
            newMetadataDescription={this.state.newMetadataDescription}
            newMetadataType={this.state.newMetadataType}
            metadataTypeOptions={metadataTypeOptions}
            choices={this.state.choices}
            choiceGroupOptions={choiceGroupOptions}
          />

          <CreateFolderModal
            isOpen={this.state.showCreateFolderModal}
            onDismiss={() => this.setState({ showCreateFolderModal: false })}
            handleCreateFolderSubmit={this.handleCreateFolderSubmit}
            columnFields={columnFields}
          />
        </div>
      </section>
    );
  }
}
