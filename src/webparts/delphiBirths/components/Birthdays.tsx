import * as React from 'react';
import styles from './Birthdays.module.scss';
import { IBirthdaysProps } from './IBirthdaysProps';
import { HappyBirthday, IUser } from '../../../controls/happybirthday';
import * as moment from 'moment';
import { IBirthdayState } from './IBirthdaysState';
import SPService from '../../../services/SPService';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
const imgBackgroundBallons: string = require('../../../../assets/ballonsBackgroud.png');
import { Image, ImageFit } from 'office-ui-fabric-react/lib/Image';
import { Label } from 'office-ui-fabric-react/lib/Label';
import * as strings from 'DelphiBirthsWebPartStrings';

export default class Birthdays extends React.Component<IBirthdaysProps, IBirthdayState> {
  private _users: IUser[] = [];
  private _spServices: SPService;
  constructor(props: IBirthdaysProps) {
    super(props);
    this._spServices = new SPService(this.props.context);
    this.state = {
      Users: [],
      showBirthdays: true
    };
  }

  public async componentDidMount(): Promise<void> {
    setTimeout(async () => {
      await this.GetUsers();
    }, 1000);
  }

  public componentDidUpdate(prevProps: IBirthdaysProps, prevState: IBirthdayState): void {

  }
  // Render
  public render(): React.ReactElement<IBirthdaysProps> {
    const _center: any = !this.state.showBirthdays ? "center" : "";
    return (
      <div className={styles.happyBirthday}
        style={{ textAlign: _center }} >
        <div className={styles.container}>
          <WebPartTitle displayMode={this.props.displayMode}
            title={this.props.title}
            updateProperty={this.props.updateProperty} />
          {
            !this.state.showBirthdays ?
              <div className={styles.backgroundImgBallons}>
                <Image imageFit={ImageFit.cover}
                  src={imgBackgroundBallons}
                  width={150}
                  height={150}
                />
                <Label className={styles.subTitle}>{this.props.MessageNoBirthdays ?? strings.MessageNoBirthdays}</Label>
              </div>
              :
              <HappyBirthday users={this.state.Users} imageTemplate={this.props.imageTemplate}
              height={this.props.height} width={this.props.width}
              />
          }
        </div>
      </div>
    );
  }

  // Sort Array of Birthdays
  private SortBirthdays(users: IUser[]): IUser[] {
    return users.sort((a, b) => {
      if (a.birthday > b.birthday) {
        return 1;
      }
      if (a.birthday < b.birthday) {
        return -1;
      }
      return 0;
    });
  }
  // Load List Of Users
  private async GetUsers(): Promise<void> {
    let _otherMonthsBirthdays: IUser[], _dezemberBirthdays: IUser[];
    const listItems = await this._spServices.getPBirthdays(this.props.numberUpcomingDays);
    if (listItems && listItems.length > 0) {
      _otherMonthsBirthdays = [];
      _dezemberBirthdays = [];
      for (const item of listItems) {
        this._users.push({
          key: item.UserName.EMail, userName: item.UserName.Title,
          message: item.message, anniversary: item.anniversary, userEmail: item.UserName.EMail,
          jobDescription: item.jobTitle,
          birthday: moment.utc(item.Birthday).local().format()
        });
      }
      // Sort Items by Birthday MSGraph List Items API don't support ODATA orderBy
      // for end of year teste and sorting
      //  first select all bithdays of Dezember to sort this must be the first to show
      if (moment().format('MM') === '12') {
        _dezemberBirthdays = this._users.filter((v) => {
          const _currentMonth = moment(v.birthday, ["MM-DD-YYYY", "YYYY-MM-DD", "DD/MM/YYYY", "MM/DD/YYYY"]).format('MM');
          return (_currentMonth === '12');
        });
        // Sort by birthday date in Dezember month
        _dezemberBirthdays = this.SortBirthdays(_dezemberBirthdays);
        // select birthdays != of month 12
        _otherMonthsBirthdays = this._users.filter((v) => {
          const _currentMonth = moment(v.birthday, ["MM-DD-YYYY", "YYYY-MM-DD", "DD/MM/YYYY", "MM/DD/YYYY"]).format('MM');
          return (_currentMonth !== '12');
        });
        // sort by birthday date
        _otherMonthsBirthdays = this.SortBirthdays(_otherMonthsBirthdays);
        // Join the 2 arrays
        this._users = _dezemberBirthdays.concat(_otherMonthsBirthdays);
      }
      else {
        this._users = this.SortBirthdays(this._users);
      }
    }

    //  this._users=[];
    this.setState(
      {
        Users: this._users,
        showBirthdays: this._users.length === 0 ? false : true
      });
  }
}
