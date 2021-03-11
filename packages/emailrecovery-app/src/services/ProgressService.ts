export interface IProgressService {
  reportProgress(activity: string, status: string): void;
  reportComplete(activity: string): void;
}

export class ProgressService implements IProgressService {

  reportProgress(activity: string, status: string): void {
    // TODO: Implement
  }

  reportComplete(activity: string): void {
    // TODO: Implement
  }
}