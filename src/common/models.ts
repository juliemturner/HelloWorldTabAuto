export interface IFile {
  Id: number;
  Title: string;
  Name: string;
  Size: number;
}

export interface IDocumentsByType {
  [key: string]: IFile[];
}