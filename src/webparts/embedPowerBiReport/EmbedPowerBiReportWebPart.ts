import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, ServiceScope } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'EmbedPowerBiReportWebPartStrings';

import EmbedPowerBiReport from './components/EmbedPowerBiReport';
import { IEmbedPowerBiReportProps } from './components/IEmbedPowerBiReportProps';
import { PowerBiWorkspace, PowerBiReport } from './../../models/PowerBiModels';
import { PowerBiService } from './../../services/PowerBiService';

export interface IEmbedPowerBiReportWebPartProps {
  workspaceId: string;
  reportId: string;
  widthToHeight: number;
}

export default class EmbedPowerBiReportWebPart extends BaseClientSideWebPart<IEmbedPowerBiReportWebPartProps> {

  private powerBiReactReport: EmbedPowerBiReport;

  private workspaceOptions: IPropertyPaneDropdownOption[];
  private workspacesFetched: boolean = false;

  private fetchWorkspaceOptions(): Promise<IPropertyPaneDropdownOption[]> {
    return PowerBiService.GetWorkspaces(this.context.serviceScope).then((workspaces: PowerBiWorkspace[]) => {
      var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
      workspaces.map((workspace: PowerBiWorkspace) => {
        options.push({ key: workspace.id, text: workspace.name });
      });
      return options;
    });
  }

  private reportOptions: IPropertyPaneDropdownOption[];
  private reportsFetched: boolean = false;

  private fetchReportOptions(): Promise<IPropertyPaneDropdownOption[]> {
    return PowerBiService.GetReports(this.context.serviceScope, this.properties.workspaceId).then((reports: PowerBiReport[]) => {
      var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
      reports.map((report: PowerBiReport) => {
        options.push({ key: report.id, text: report.name });
      });
      return options;
    });
  }


  public render(): void {
    console.log("EmbedPowerBiReportWebPart.render");
    const element: React.ReactElement<IEmbedPowerBiReportProps> = React.createElement(
      EmbedPowerBiReport,
      {
        webPartContext: this.context,
        serviceScope: this.context.serviceScope,
        defaultWorkspaceId: this.properties.workspaceId,
        defaultReportId: this.properties.reportId,
        defaultWidthToHeight: this.properties.widthToHeight
      }
    );

    this.powerBiReactReport = <EmbedPowerBiReport>ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    console.log("onPropertyPaneConfigurationStart");
    if (this.workspacesFetched && this.reportsFetched) {
      return;
    }

    if (this.workspacesFetched && !this.reportsFetched) {
      this.powerBiReactReport.setState({ loading: true });
      this.fetchReportOptions().then((options: IPropertyPaneDropdownOption[]) => {
        this.reportOptions = options;
        this.reportsFetched = true;
        this.powerBiReactReport.setState({ loading: false });
        this.context.propertyPane.refresh();
        this.render();
      });
      return;
    }

    this.powerBiReactReport.setState({ loading: true });
    this.fetchWorkspaceOptions().then((options: IPropertyPaneDropdownOption[]) => {
      this.workspaceOptions = options;
      this.workspacesFetched = true;
      this.powerBiReactReport.setState({ loading: false });
      this.context.propertyPane.refresh();
      this.render();
    });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    console.log("onPropertyPaneFieldChanged");
    if (propertyPath === 'workspaceId' && newValue) {
      console.log("Workspace ID updated: " + newValue);
      // reset report settings
      this.properties.reportId = "";
      this.reportOptions = [];
      this.reportsFetched = false;
      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();
      // communicate loading items      
      this.powerBiReactReport.setState({ loading: true, workspaceId: this.properties.workspaceId });
      this.fetchReportOptions().then((options: IPropertyPaneDropdownOption[]) => {
        this.reportOptions = options;
        this.reportsFetched = true;
        this.powerBiReactReport.setState({ loading: false });
        this.context.propertyPane.refresh();
      });
    }

    if (propertyPath === 'reportId' && newValue) {
      this.powerBiReactReport.setState({ reportId: this.properties.reportId });
    }

    if (propertyPath === 'widthToHeight' && newValue) {
      this.powerBiReactReport.setState({ widthToHeight: this.properties.widthToHeight });
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    console.log("getPropertyPaneConfiguration");
    return {
      pages: [
        {
          header: {
            description: "A gratuitous demo of embedding Power BI reports using a React Web Part"
          },
          groups: [
            {
              groupName: "Power BI Configuration",
              groupFields: [
                PropertyPaneDropdown(
                  "workspaceId", {
                    label: "Select a Workspace",
                    options: this.workspaceOptions,
                    disabled: !this.workspacesFetched
                  }),
                PropertyPaneDropdown(
                  "reportId", {
                    label: "Select a Report",
                    options: this.reportOptions,
                    disabled: !this.reportsFetched
                  }),
                  PropertyPaneSlider("widthToHeight", {
                    label: "Width to Height Percentage",
                    min: 25,
                    max: 400
                  })          
                ]
            }
          ]
        }
      ]
    };
  }
}
