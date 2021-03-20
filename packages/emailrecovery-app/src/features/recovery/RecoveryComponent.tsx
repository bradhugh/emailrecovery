import {
  ActionButton,
  ChoiceGroup,
  IChoiceGroupOption,
  Modal,
  PrimaryButton,
  TextField,
} from "@fluentui/react";
import { useId } from "@uifabric/react-hooks";
import React, { useEffect } from "react";
import { useDispatch, useSelector } from "react-redux";
import { RootState } from "../../app/store";
import { ProgressComponent } from "../progress/ProgressComponent";
import {
  loadFolderHierarchyAsync,
  setFolderName,
  setFolderDialogOpen,
  setSourceFolder,
  promptForFolderNameAsync,
  performRecoveryAsync,
} from "./recoverySlice";

const options: IChoiceGroupOption[] = [
  { key: "recoverableitemsdeletions", text: "Recoverable Items" },
  { key: "recoverableitemspurges", text: "Purges" },
];

const FolderNameComponent: React.FC = () => {
  const titleId = useId("title");
  const { folderName, isFolderDialogOpen } = useSelector(
    (state: RootState) => state.recovery
  );
  const dispatch = useDispatch();
  return (
    <Modal titleAriaId={titleId} isOpen={isFolderDialogOpen} isBlocking={true}>
      <div>
        <span id={titleId}>Folder name</span>
      </div>
      <div>
        <TextField
          label="The restore process will create a folder for the items. Enter a name for the folder or you can accept the default."
          value={folderName}
          onChange={(_, val) => dispatch(setFolderName(val ?? ""))}
        />
      </div>
      <div>
        <PrimaryButton
          onClick={() => {
            dispatch(setFolderDialogOpen(false));
            dispatch(performRecoveryAsync());
          }}
        >
          Create Folder
        </PrimaryButton>
        <ActionButton
          onClick={() => {
            dispatch(setFolderName(""));
            dispatch(setFolderDialogOpen(false));
          }}
        >
          Cancel
        </ActionButton>
      </div>
    </Modal>
  );
};

export const RecoveryComponent: React.FC = () => {
  const dispatch = useDispatch();
  const { sourceFolder } = useSelector((state: RootState) => state.recovery);
  useEffect(() => void dispatch(loadFolderHierarchyAsync()), [dispatch]);

  return (
    <>
      <ChoiceGroup
        options={options}
        required={true}
        selectedKey={sourceFolder}
        onChange={(e, opt) =>
          opt?.key ? dispatch(setSourceFolder(opt.key)) : null
        }
      />
      <PrimaryButton onClick={() => dispatch(promptForFolderNameAsync())}>
        Start Recovery
      </PrimaryButton>
      <FolderNameComponent />
      <ProgressComponent />
    </>
  );
};
