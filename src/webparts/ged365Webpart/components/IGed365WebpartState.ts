import { IDropdownOption } from "office-ui-fabric-react";
import { SPListItem, SPListColumn} from "../../Services/SPServices";


export interface IGed365WebpartState {
  listColumns: SPListColumn[];
  listTiltes: IDropdownOption[];
  listItems: SPListItem[];
  status: string;
  Titre_list_item: string;
  showModal: boolean;
  listItemId: string;
  selectedDocumentType: string;
  metadata: { [key: string]: any };
  uploadFile: File | null;
  isUploadMode: boolean;
  showCreateModal: boolean;
  showUploadModal: boolean;
  showAddMetadataModal: boolean;
  newMetadataField: string;
  newMetadataDescription: string;
  newMetadataType: string;
  choices: string[];
}





