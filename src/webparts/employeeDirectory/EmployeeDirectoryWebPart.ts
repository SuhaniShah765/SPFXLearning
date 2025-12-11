import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import EmployeeDirectory from './components/EmployeeDirectory';
import { IEmployeeDirectoryProps } from './components/IEmployeeDirectoryProps';

export interface IEmployeeDirectoryWebPartProps {
  listName: string;
}

export default class EmployeeDirectoryWebPart extends BaseClientSideWebPart<IEmployeeDirectoryWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IEmployeeDirectoryProps> = React.createElement(
      EmployeeDirectory,
      {
        context: this.context,
        listName: this.properties.listName || 'Employee Directory'
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
