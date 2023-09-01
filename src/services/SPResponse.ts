export interface SPResponse {
    "@odata.context": string
    value: SPResponseValue[]
  }
  
  export interface SPResponseValue {
    "@odata.type": string
    "@odata.id": string
    "@odata.etag": string
    "@odata.editLink": string
    Birthday: string
    "UserName@odata.navigationLink": string
    UserName: UserName
  }
  
  export interface UserName {
    "@odata.type": string
    "@odata.id": string
    Title: string
    EMail: string
  }
  