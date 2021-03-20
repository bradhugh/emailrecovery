import { createSlice, PayloadAction } from "@reduxjs/toolkit";
import { AppThunk } from "../../app/store";
import { EwsService } from "../../services/EwsService";
import { FolderHierarchy } from "../../services/FolderHierarchy";
import { IFolder } from "../../services/IFolder";
import { ItemCopier } from "../../services/ItemCopier";
import { Strings } from "../../Strings";
import { Utils } from "../../Utils";
import { reportProgress, reportComplete } from "../progress/progressSlice";

interface IRecoveryState {
  folders: IFolder[];
  errorMessage: string;
  sourceFolder: string;
  folderName: string;
  isFolderDialogOpen: boolean;
}

const initialState: IRecoveryState = {
  folders: [],
  errorMessage: "",
  sourceFolder: "recoverableitemsdeletions",
  folderName: "",
  isFolderDialogOpen: false,
};

export const recoverySlice = createSlice({
  name: "recovery",
  initialState,
  reducers: {
    setFolders: (state, action: PayloadAction<IFolder[]>) => {
      state.folders = action.payload;
    },
    setError: (state, action: PayloadAction<string>) => {
      state.errorMessage = action.payload;
    },
    setFolderName: (state, action: PayloadAction<string>) => {
      state.folderName = action.payload;
    },
    setFolderDialogOpen: (state, action: PayloadAction<boolean>) => {
      state.isFolderDialogOpen = action.payload;
    },
    setSourceFolder: (state, action: PayloadAction<string>) => {
      state.sourceFolder = action.payload;
    }
  },
});

export const { setFolders, setError, setFolderName, setFolderDialogOpen, setSourceFolder } = recoverySlice.actions;

export const loadFolderHierarchyAsync = (): AppThunk => async (dispatch) => {
  // TODO: Inject EwsService
  const hierarchy = new FolderHierarchy(EwsService.Default);

  dispatch(
    reportProgress({
      activity: "Initializing",
      status: "Loading folder hierarchy",
    })
  );

  await hierarchy.initialize();

  dispatch(
    reportProgress({
      activity: "Initializing",
      status: "Initialization complete",
    })
  );
  await Utils.setTimeoutAsync(1000);
  dispatch(reportComplete("Initializing"));

  dispatch(setFolders(hierarchy.folders));
};

export const reportError = (activity: string, error: Error): AppThunk => (
  dispatch
) => {
  dispatch(reportComplete(activity));
  dispatch(setError(error.message));
};

const createFolderAsync = async (folderName: string): Promise<string> => {
  const res = await EwsService.Default.createFolderAsync(
    "msgfolderroot",
    folderName
  );
  if (res.responseClass === "Success") {
    return res.folderId;
  } else {
    throw new Error(res.responseCode);
  }
};

export const promptForFolderNameAsync = (): AppThunk => async dispatch => {
  dispatch(setFolderName("Email Recovery " + new Date(Date.now()).toISOString()));
  dispatch(setFolderDialogOpen(true));
};

export const performRecoveryAsync = (): AppThunk => async (
  dispatch,
  getState,
  services
) => {
  const copyChunkSize = 50;

  const { recovery: { sourceFolder, folderName } } = getState();
  if (!sourceFolder || !folderName) {
    return;
  }

  dispatch(
    reportProgress({
      activity: Strings.recoveryInProgress,
      status: "Creating target folder",
    })
  );

  let targetFolder: string;
  try {
    targetFolder = await createFolderAsync(folderName);
    dispatch(
      reportProgress({
        activity: Strings.recoveryInProgress,
        status: "Folder created",
      })
    );
  } catch (error) {
    return reportError(Strings.recoveryInProgress, error);
  }

  const { recovery: { folders } } = getState();
  const hierarchy = new FolderHierarchy(EwsService.Default, folders);

  const progressCallback = (status: string) => {
    dispatch(reportProgress({ activity: Strings.recoveryInProgress, status }));
  };

  const copier = new ItemCopier(
    EwsService.Default,
    hierarchy,
    progressCallback,
    sourceFolder,
    targetFolder,
    copyChunkSize
  );

  let completed = false;
  do {
    try {
      completed = await copier.process();
    } catch (error) {
      reportError(Strings.recoveryInProgress, error);
    }
  } while (!completed);

  dispatch(reportProgress({ activity: Strings.recoveryInProgress, status: "Recovery completed" }));
  await Utils.setTimeoutAsync(1000);
  dispatch(reportComplete(Strings.recoveryInProgress));

  // TODO: Success message
};

export default recoverySlice.reducer;
