import * as React from 'react';
import { Button, Modal, Dropdown, IDropdownOption, ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react';
import styles from './Ged365Webpart.module.scss';
import { IGed365WebpartProps } from './IGed365WebpartProps';
import { IGed365WebpartState } from './IGed365WebpartState';
import { SPOperations, SPListColumn } from '../../Services/SPServices';

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
      Titre_list_item: '', // Initialisation de Titre_list_item
      showModal: false,
      listItemId: '',
      selectedDocumentType: "txt",
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
              metadata
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
        [internalName]: value
      }
    }));
  };

  private handleSelectChange = (internalName: string) => (event: React.ChangeEvent<HTMLSelectElement>) => {
    const value = event.target.value;
    this.setState(prevState => ({
      metadata: {
        ...prevState.metadata,
        [internalName]: value
      }
    }));
  };

  private handleTextareaChange = (internalName: string) => (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    const value = event.target.value;
    this.setState(prevState => ({
      metadata: {
        ...prevState.metadata,
        [internalName]: value
      }
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
                choices: ['']
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

    this.setState({ showCreateModal: false, selectedDocumentType: "txt" });
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
      color: textColor // Set the text color for the buttons
    };

    const documentTypeOptions: IDropdownOption[] = [
      { key: 'docx', text: 'Word Document (.docx)' },
      { key: 'txt', text: 'Text Document (.txt)' },
      { key: 'pptx', text: 'PowerPoint Presentation (.pptx)' },
      { key: 'xlsx', text: 'Excel Spreadsheet (.xlsx)' }
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
                        {column.choices && column.choices.map(choice => (
                          <option key={choice} value={choice}>{choice}</option>
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
          <div className={styles['align-right-items']}>
            <Button 
              text="+ Créer un document" 
              onClick={this.openCreateModal} 
              className={`${this.getButtonClass()} ${styles.myButton}`} 
              style={buttonStyle} // Apply the style to the button
            />
            <Button 
              text="+ Ajouter un document" 
              onClick={this.openUploadModal} 
              className={`${this.getButtonClass()} ${styles.myButton}`} 
              style={buttonStyle} // Apply the style to the button
            />
            <Button 
              text="+ Ajouter une métadonnée" 
              onClick={this.openAddMetadataModal} 
              className={`${this.getButtonClass()} ${styles.myButton}`} 
              style={buttonStyle} // Apply the style to the button
              disabled={!this.props.list_title} // Disable if list_title is not selected
            />
          </div>

          <Modal
            isOpen={this.state.showCreateModal}
            onDismiss={() => this.setState({ showCreateModal: false })}
            isBlocking={false}
            containerClassName={styles.modalContainer}
          >
            <div className={styles.modalHeader}>
              <span>Créer un document</span>
              <Button iconProps={{ iconName: 'Cancel' }} onClick={() => this.setState({ showCreateModal: false })} />
            </div>
            <div className={styles.modalBody}>
              {columnFields}

              <div className="mb-3">
                <label>Select Document Type</label>
                <div className={styles['field-wrapper']}>
                  <div className={styles['field-group']}>
                    <Dropdown
                      placeholder="Select document type"
                      options={documentTypeOptions}
                      onChange={this.handleDocumentTypeChange}
                    />
                  </div>
                </div>
              </div>
            </div>
            <div className={styles.modalFooter}>
              <Button className={styles.myButton} text="Submit" onClick={this.handleCreateSubmit} />
              <Button className={styles.cancelButton} text="Cancel" onClick={() => this.setState({ showCreateModal: false })} />
            </div>
          </Modal>

          <Modal
            isOpen={this.state.showUploadModal}
            onDismiss={() => this.setState({ showUploadModal: false })}
            isBlocking={false}
            containerClassName={styles.modalContainer}
          >
            <div className={styles.modalHeader}>
              <span>+ Ajouter un document</span>
              <Button iconProps={{ iconName: 'Cancel' }} onClick={() => this.setState({ showUploadModal: false })} />
            </div>
            <div className={styles.modalBody}>
              <div className="mb-3">
                <label htmlFor="uploadFile">Select File to Upload</label>
                <div className={styles['field-wrapper']}>
                  <div className={styles['field-group']}>
                    <input
                      type="file"
                      id="uploadFile"
                      className={styles['text-field']}
                      onChange={this.handleFileChange}
                    />
                  </div>
                </div>
              </div>

              {columnFields}
            </div>
            <div className={styles.modalFooter}>
              <Button className={styles.myButton} text="Submit" onClick={this.handleUploadSubmit} />
              <Button className={styles.cancelButton} text="Cancel" onClick={() => this.setState({ showUploadModal: false })} />
            </div>
          </Modal>

          <Modal
            isOpen={this.state.showAddMetadataModal}
            onDismiss={() => this.setState({ showAddMetadataModal: false })}
            isBlocking={false}
            containerClassName={styles.modalContainer}
          >
            <div className={styles.modalHeader}>
              <span>+ Ajouter une métadonnée</span>
              <Button iconProps={{ iconName: 'Cancel' }} onClick={() => this.setState({ showAddMetadataModal: false })} />
            </div>
            <div className={styles.modalBody}>
              <div className="mb-3">
                <label htmlFor="newMetadataField">Name</label>
                <div className={styles['field-wrapper']}>
                  <div className={styles['field-group']}>
                    <input
                      type="text"
                      id="newMetadataField"
                      className={styles['text-field']}
                      onChange={this.handleNewMetadataFieldChange}
                      value={this.state.newMetadataField}
                    />
                  </div>
                </div>
              </div>
              <div className="mb-3">
                <label htmlFor="newMetadataDescription">Description</label>
                <div className={styles['field-wrapper']}>
                  <div className={styles['field-group']}>
                    <input
                      type="text"
                      id="newMetadataDescription"
                      className={styles['text-field']}
                      onChange={this.handleNewMetadataDescriptionChange}
                      value={this.state.newMetadataDescription}
                    />
                  </div>
                </div>
              </div>
              <div className="mb-3">
                <label htmlFor="newMetadataType">Metadata Type</label>
                <div className={styles['field-wrapper']}>
                  <div className={styles['field-group']}>
                    <Dropdown
                      placeholder="Select metadata type"
                      options={metadataTypeOptions}
                      onChange={this.handleNewMetadataTypeChange}
                      selectedKey={this.state.newMetadataType}
                    />
                  </div>
                </div>
              </div>
              {this.state.newMetadataType === 'Text' && (
                <div className="mb-3">
                  <label>Text Value</label>
                  <div className={styles['field-wrapper']}>
                    <div className={styles['field-group']}>
                      <input
                        type="text"
                        className={styles['text-field']}
                        value=""
                        readOnly
                      />
                    </div>
                  </div>
                </div>
              )}
              {this.state.newMetadataType === 'Choice' && (
                <>
                  {this.state.choices.map((choice, index) => (
                    <div key={index} className="mb-3">
                      <label>Choice {index + 1}</label>
                      <div className={styles['field-wrapper']}>
                        <div className={styles['field-group']}>
                          <input
                            type="text"
                            className={styles['text-field']}
                            value={choice}
                            onChange={this.handleChoiceChange(index)}
                          />
                        </div>
                      </div>
                    </div>
                  ))}
                  <Button text="Add Choice" onClick={this.addChoiceField} />
                </>
              )}
              {this.state.newMetadataType === 'Number' && (
                <div className="mb-3">
                  <label>Number Value</label>
                  <div className={styles['field-wrapper']}>
                    <div className={styles['field-group']}>
                      <input
                        type="number"
                        className={styles['text-field']}
                        value=""
                        readOnly
                      />
                    </div>
                  </div>
                </div>
              )}
              {this.state.newMetadataType === 'Boolean' && (
                <div className="mb-3">
                  <label>Boolean Value</label>
                  <div className={styles['field-wrapper']}>
                    <div className={styles['field-group']}>
                      <ChoiceGroup options={choiceGroupOptions} defaultSelectedKey="No" />
                    </div>
                  </div>
                </div>
              )}
              {this.state.newMetadataType === 'Image' && (
                <div className="mb-3">
                  <label>Upload Image</label>
                  <div className={styles['field-wrapper']}>
                    <div className={styles['field-group']}>
                      <input
                        type="file"
                        accept="image/*"
                        className={styles['text-field']}
                      />
                    </div>
                  </div>
                </div>
              )}
            </div>
            <div className={styles.modalFooter}>
              <Button className={styles.myButton} text="Add" onClick={this.addMetadataField} />
              <Button className={styles.cancelButton} text="Cancel" onClick={() => this.setState({ showAddMetadataModal: false })} />
            </div>
          </Modal>
        </div>
      </section>
    );
  }
}
