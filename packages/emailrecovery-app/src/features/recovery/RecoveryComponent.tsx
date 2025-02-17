import {
  ChoiceGroup,
  DefaultButton,
  FontWeights,
  getTheme,
  IChoiceGroupOption,
  mergeStyleSets,
  MessageBar,
  MessageBarType,
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

const theme = getTheme();
const contentStyles = mergeStyleSets({
  container: {
    display: "flex",
    flexFlow: "column nowrap",
    alignItems: "stretch",
  },
  header: [
    theme.fonts.xLargePlus,
    {
      flex: "1 1 auto",
      borderTop: `4px solid ${theme.palette.themePrimary}`,
      color: theme.palette.neutralPrimary,
      display: "flex",
      alignItems: "center",
      fontWeight: FontWeights.semibold,
      padding: "12px 12px 14px 24px",
    },
  ],
  body: {
    flex: "4 4 auto",
    padding: "0 24px 24px 24px",
    overflowY: "hidden",
    selectors: {
      p: { margin: "14px 0" },
      "p:first-child": { marginTop: 0 },
      "p:last-child": { marginBottom: 0 },
    },
  },
  buttonBar: {
    flex: "1 1 auto",
    display: "flex",
    padding: "12px 12px 14px 24px",
  },
});

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
    <Modal
      titleAriaId={titleId}
      isOpen={isFolderDialogOpen}
      isBlocking={true}
      containerClassName={contentStyles.container}
      styles={{ root: { alignItems: "flex-start", paddingTop: 14 } }}
    >
      <div className={contentStyles.header}>
        <span id={titleId}>Folder name</span>
      </div>
      <div className={contentStyles.body}>
        <TextField
          label="The restore process will create a folder for the items. Enter a name for the folder or you can accept the default."
          value={folderName}
          onChange={(_, val) => dispatch(setFolderName(val ?? ""))}
        />
      </div>
      <div className={contentStyles.buttonBar}>
        <PrimaryButton
          onClick={() => {
            dispatch(setFolderDialogOpen(false));
            dispatch(performRecoveryAsync());
          }}
          styles={{ root: { margin: 5 } }}
        >
          Create Folder
        </PrimaryButton>
        <DefaultButton
          onClick={() => {
            dispatch(setFolderName(""));
            dispatch(setFolderDialogOpen(false));
          }}
          styles={{ root: { margin: 5 } }}
        >
          Cancel
        </DefaultButton>
      </div>
    </Modal>
  );
};

var docUrl = "https://support.microsoft.com/office/recover-and-restore-deleted-items-in-outlook-49e81f3c-c8f4-4426-a0b9-c0fd751d48ce";
var decommDate = new Date(2024, 2, 17);
export const RecoveryComponent: React.FC = () => {
  const dispatch = useDispatch();
  const { sourceFolder } = useSelector((state: RootState) => state.recovery);
  useEffect(() => void dispatch(loadFolderHierarchyAsync()), [dispatch]);

  return (
    <>
      <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>
        Email Recovery Add-in is now decommissioned as of {decommDate.toLocaleDateString()}. Use the process documented in the 
        <a href={docUrl} target="_blank" rel="noreferrer">Recover and restore deleted items in Outlook</a> instead.
      </MessageBar>
      {/* <div style={{ margin: 5 }}>
        <ChoiceGroup
          label="Select the folder from which you wish to recover email"
          options={options}
          required={true}
          selectedKey={sourceFolder}
          onChange={(e, opt) =>
            opt?.key ? dispatch(setSourceFolder(opt.key)) : null
          }
        />
        <PrimaryButton
          onClick={() => dispatch(promptForFolderNameAsync())}
          styles={{ root: { marginTop: 10 } }}
        >
          Start Recovery
        </PrimaryButton>
        <FolderNameComponent />
        <ProgressComponent />
      </div> */}
    </>
  );
};
