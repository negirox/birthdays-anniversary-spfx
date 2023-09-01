import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as moment from 'moment';
import { Log } from "@microsoft/sp-core-library";
import { IUserProperties } from "./IUserProperties";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views/list";
import { getSP } from "../pnpjsConfig";
import {
  SPHttpClient
} from '@microsoft/sp-http';
import { SPResponse } from "./SPResponse";
const LOG_SOURCE = "MSGraphService";

export class SPService {
  private graphClient: MSGraphClientV3 = null;
  private birthdayListTitle: string = "Birthdays";
  private _context: WebPartContext
  constructor(context: WebPartContext) {
    this._context = context;
    this.InitGraph();
  }
  private async InitGraph(): Promise<void> {
    this.graphClient = await this._context.msGraphClientFactory.getClient('3');
  }

  public async getUserPropertiesByLastName(searchFor: string, client: MSGraphClientV3): Promise<IUserProperties[]> {
    const userProperties: IUserProperties[] = [];
    try {
      const res = await client.api("users/delta")
        .version("v1.0").get();
      if (res.value.length !== 0) {
        // eslint-disable-next-line @typescript-eslint/no-unused-vars, @typescript-eslint/no-explicit-any
        res.value.map((_userProperty: any, _index: any) => {
          if (_userProperty.mail !== null) {
            userProperties.push({
              businessPhone: _userProperty.businessPhones[0],
              displayName: _userProperty.displayName,
              email: _userProperty.mail,
              JobTitle: _userProperty.jobTitle,
              OfficeLocation: _userProperty.officeLocation,
              mobilePhone: _userProperty.mobilePhone,
              preferredLanguage: _userProperty.preferredLanguage
            });
          }
        });
      }
    } catch (error) {
      console.log(error);
      Log.error(LOG_SOURCE + "getUserPropertiesByLastName():", error);
    }
    return userProperties;

  }
  public async GetUsersBirthday(userId: string): Promise<any> {
    const url = `users/${userId}/birthday`;
    const res = await this.graphClient.api(url).version("v1.0").get();
    return res;
  }
  public async getAllUsers(uri: string): Promise<any> {
    try {
      if (this.graphClient === null) {
        await this.InitGraph();
      }
      // get users 'https://graph.microsoft.com/v1.0/users/delta?$select=displayName,jobTitle,mail,Id'; or  nextLink URL
      const _users = await this.graphClient.api(uri).version("v1.0").get();
      // has data?
      if (_users && _users.value && _users.value.length === 0) {
        return;
      }
      // get deltaLink for track changes.
      // get nextLink to get next page
      const _nextLink = (typeof _users["@odata.nextLink"] !== undefined) ? _users["@odata.nextLink"] : undefined;
      const _deltaLink = (typeof _users["@odata.deltaLink"] !== undefined) ? _users["@odata.deltaLink"] : undefined;
      // Read Users
      for (const user of _users.value) {
        // If user was removed from AAD
        try {
          /* if (user['@removed']) {
              await deleteUser(user);
              continue;
          } */
          const _birthday = await this.GetUsersBirthday(user.id);
          console.log(_birthday);
          //const _year = moment(_birthday.toString()).format('YYYY');
          // The Birthday Date has year 2000
          /*  if (_year === '2000') {
               // check if user exists
               _exists = await checkUserExist(user);
               if (!_exists) {
                   // Add user to List
                   await addUser(user, _birthday)
               } else {
                   //Update user
                   await updateUser(user, _birthday)
               }
           } */
        } catch (error) {
          console.log(`Error Adding or Updating users : ${error} `);
        }
      }
      try {
        // Load next Page
        if (_nextLink !== ('' || undefined || null)) {
          await this.getAllUsers(_nextLink);
        }
        // deltaLink exist (last request)
        if (_deltaLink) {
          // Save Tenant property with deltaLink for track changes
          console.log('Last request');
          console.log(_deltaLink);

        }

      } catch (error) {
        console.log(`Error updating StorageEntity : ${error} `);
      }

    } catch (error) {
      console.log(`Error on read users : ${error} `);
    }

    return;
  }
  public async ensureBirthdaysList(): Promise<boolean> {
    const sp = getSP(this._context);
    const _web = sp.web;
    let result = false;

    try {
      await sp.web.lists.getByTitle(this.birthdayListTitle)().then(x => {
        console.log('list already exist');
      }).catch(async x => {

        const ensureResult = await _web.lists.ensure(this.birthdayListTitle, "Birthdays And Anniversary", 100, true);
        // if we've got the list
        if (ensureResult.list !== null) {

          // if the list has just been created
          if (ensureResult.created) {
            // we need to add the custom fields to the list
            const jobTitleFieldAddResult = await sp.web.lists.getByTitle(this.birthdayListTitle).fields.addText(
              "jobTitle", { MaxLength: 255, Required: false });
            await jobTitleFieldAddResult.field.update({ Title: "Job Title" });
            const emailFieldAddResult = await ensureResult.list.fields.addText(
              "email", { MaxLength: 255, Required: false });
            await emailFieldAddResult.field.update({ Title: "Email" });
            const BirthdayFieldAddResult = await ensureResult.list.fields.addDateTime(
              "Birthday",
            );
            await BirthdayFieldAddResult.field.update({ Title: "Birthday" });
            const userAADGUIDFieldAddResult = await ensureResult.list.fields.addText(
              "userAADGUID", { MaxLength: 255, Required: false });
            await userAADGUIDFieldAddResult.field.update({ Title: "AAD ID " });
            await ensureResult.list.fields.getByInternalNameOrTitle('Title').update({ Title: 'Display Name' });
            // the list is ready to be used
            await ensureResult.list.fields.addUser('UserName', { Description: "User Name", Required: true });
            await ensureResult.list.fields.addText("message", { MaxLength: 255, Required: false });
            await ensureResult.list.fields.addBoolean("anniversary", { Required: false });
            const allItemsView = ensureResult.list.views.getByTitle('All Items');
            await allItemsView.fields.add('UserName');
            await allItemsView.fields.add('Birthday');
            await allItemsView.fields.add('message');
            await allItemsView.fields.add('anniversary');
            result = true;
          }
        }
      });

    } catch (e) {
      // if we fail to create the list, write an exception in the _context log
      result = false;
    }

    return result;
  }
  // Get Profiles
  public async getPBirthdays(upcommingDays: number): Promise<any[]> {
    let _results: SPResponse, _today: string, _month: number, _day: number;
    let _filter: string, _countdays: number, _f: number, _nextYearStart: string;
    let _FinalDate: string;
    try {
      _results = null;
      const _currentYear = new Date().getFullYear().toString();
      _today = `${_currentYear}-${moment().format('MM-DD')}`;
      _month = parseInt(moment().format('MM'));
      _day = parseInt(moment().format('DD'));
      _filter = "Birthday ge '" + _today + "'";
      // If we are in December we have to look if there are birthdays in January
      // we have to build a condition to select birthday in January based on number of upcommingDays
      // we can not use the year for test, the year is always 2000.
      console.log(_month);
      _countdays = _day + upcommingDays;
      _f = 0;
      if (_month === 12 && _countdays > 31) {
        _nextYearStart = `${_currentYear}-01-01`;
        _FinalDate = `${_currentYear}-01-`;
        _f = _countdays - 31;
        _FinalDate = _FinalDate + _f;
        _filter = "Birthday ge '" + _today + "' or (Birthday ge '" + _nextYearStart + "' and Birthday le '" + _FinalDate + "')";
      }
      else {
        _FinalDate = `${_currentYear}-`;
        if ((_countdays) > 31) {
          _f = _countdays - 31;
          _month = _month + 1;
          _FinalDate = _FinalDate + _month + '-' + _f;
        } else {
          _FinalDate = _FinalDate + _month + '-' + _countdays;
        }
        _filter = "Birthday ge '" + _today + "' and Birthday le '" + _FinalDate + "'";
      }

      this.graphClient = await this._context.msGraphClientFactory.getClient('3');
      const ConfigUrl = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.birthdayListTitle}')/Items?$expand=UserName&$select=UserName/Title,UserName/EMail,Birthday,jobTitle,*&$filter=${_filter}&$top=4999&$orderby=Birthday desc`;

      const response = await this._context.spHttpClient.get(ConfigUrl, SPHttpClient.configurations.v1);
      _results = await response.json();
      return _results.value;

    } catch (error) {
      console.dir(error);
      return Promise.reject(error);
    }
  }
}
export default SPService;
