import * as React from 'react';
import styles from './DelphiBirths.module.scss';
import { IDelphiBirthsProps } from './IDelphiBirthsProps';
import Birthdays from './Birthdays';


export default class DelphiBirths extends React.Component<IDelphiBirthsProps, {}> {
  public render(): React.ReactElement<IDelphiBirthsProps> {
    return (
      <section className={styles.delphiBirths}>
        <Birthdays context={this.props.context} displayMode={this.props.displayMode}
          imageTemplate={this.props.imageTemplate} numberUpcomingDays={this.props.numberUpcomingDays}
          title={this.props.title} updateProperty={this.props.updateProperty} children={this.props.children}
          height={this.props.height} width={this.props.width}
          MessageNoBirthdays={this.props.MessageNoBirthdays}
        />
      </section>
    );
  }
}
