import { IUser } from './IUser';

export interface IHappyBirthdayProps {
  users: IUser[];
  imageTemplate: string;
  height?:string;
  width?:string;
}
