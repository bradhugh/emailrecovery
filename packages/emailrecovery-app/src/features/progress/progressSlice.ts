import { createSlice, PayloadAction } from "@reduxjs/toolkit";

interface IProgressState {
  active: boolean;
  activity: string;
  status: string;
}

const initialState: IProgressState = {
  active: false,
  activity: "",
  status: "",
};

export const progressSlice = createSlice({
  name: "progress",
  initialState,
  reducers: {
    reportProgress: (state, action: PayloadAction<{ activity: string, status: string}>) => {
      state.active = true;
      state.activity = action.payload.activity;
      state.status = action.payload.status;
    },
    reportComplete: (state, action: PayloadAction<string>) => {
      state.activity = action.payload;
      state.active = false;
    }
  },
});

export const { reportProgress, reportComplete } = progressSlice.actions;

export default progressSlice.reducer;
