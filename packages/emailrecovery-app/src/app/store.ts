import {
  configureStore,
  ThunkAction,
  Action,
  ThunkDispatch,
} from "@reduxjs/toolkit";
import counterReducer from "../features/counter/counterSlice";
import progressReducer from "../features/progress/progressSlice";
import recoveryReducer from "../features/recovery/recoverySlice";
import { IServices } from "../services/IServices";

export const store = configureStore({
  reducer: {
    counter: counterReducer,
    recovery: recoveryReducer,
    progress: progressReducer,
  },
});

export type RootState = ReturnType<typeof store.getState>;
export type AppThunk<ReturnType = void> = ThunkAction<
  ReturnType,
  RootState,
  IServices,
  Action<string>
>;

export type AppDispatch = ThunkDispatch<RootState, IServices, Action<string>>;
