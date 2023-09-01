import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'DelphiBirthsWebPartStrings';
import DelphiBirths from './components/DelphiBirths';
import { IDelphiBirthsProps } from './components/IDelphiBirthsProps';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls';
import SPService from '../../services/SPService';

export interface IDelphiBirthsWebPartProps {
  title: string;
  numberUpcomingDays: number;
  template: any;
  height:string;
  width:string;
}

const imageTemplate: { imageUrl: string }[] = [{
  imageUrl: require('./../../../assets/cof.svg')
},
{
  imageUrl: require('./../../../assets/cof5.svg')
},
{
  imageUrl: require('./../../../assets/cof1.svg')
},
{
  imageUrl: require('./../../../assets/cof3.svg')
},
{
  imageUrl: require('./../../../assets/cof8.svg')
},
{
  imageUrl: require('./../../../assets/ballons.svg')
},
{
  imageUrl: require('./../../../assets/cof2.svg')
},
{
  imageUrl: require('./../../../assets/cof10.svg')
},
{
  imageUrl: require('./../../../assets/cof11.svg')
},
{
  imageUrl: require('./../../../assets/cof12.svg')
},
{
  imageUrl: require('./../../../assets/cof14.svg')
},
{
  imageUrl: require('./../../../assets/cof14_1.svg')
},
{
  imageUrl: require('./../../../assets/cof18.svg')
},
{
  imageUrl: require('./../../../assets/cof17.svg')
},
{
  imageUrl: require('./../../../assets/cof19.svg')
},
{
  imageUrl: require('./../../../assets/cof20.svg')
},
{
  imageUrl: require('./../../../assets/cof22.svg')
},
{
  imageUrl: require('./../../../assets/cof24.svg')
},
{
  imageUrl: require('./../../../assets/cof28.svg')
},
{
  imageUrl: require('./../../../assets/cof29.svg')
},
{
  imageUrl: require('./../../../assets/cof30.svg')
},
];
export default class DelphiBirthsWebPart extends BaseClientSideWebPart<IDelphiBirthsWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IDelphiBirthsProps> = React.createElement(
      DelphiBirths,
      {
        title: this.properties.title,
        numberUpcomingDays: this.properties.numberUpcomingDays,
        context: this.context,
        displayMode: this.displayMode,
        imageTemplate: this.properties.template,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        height:this.properties.height,
        width:this.properties.width
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public onInit(): Promise<void> {

    return super.onInit().then(async _ => {
      // other init code may be present
      const spSerrvice = new SPService(this.context);
      await spSerrvice.ensureBirthdaysList();
      //await spSerrvice.getAllUsers("/users/delta?$select=displayName,jobTitle,mail,Id&$top=999");
    });
  }
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldNumber("numberUpcomingDays", {
                  key: "numberUpcomingDays",
                  label: strings.NumberUpComingDaysLabel,
                  description: strings.NumberUpComingDaysLabel,
                  value: this.properties.numberUpcomingDays,
                  maxValue: 10,
                  minValue: 1,
                  disabled: false
                }),
                PropertyPaneChoiceGroup('template', {
                  label: 'Background Image',
                  options: imageTemplate.map((image, i) => {
                    return (
                      {
                        text: `Image ${i}`, key: i,
                        imageSrc: image.imageUrl,
                        imageSize: { width: 80, height: 80 },
                        selectedImageSrc: image.imageUrl
                      }
                    );
                  })
                }
                ),
               /*  PropertyPaneTextField('height', {
                  label: 'Enter tiles height'
                }),
                PropertyPaneTextField('width', {
                  label: 'Enter tiles width'
                }), */
              ]
            }
          ]
        }
      ]
    };
  }
}
