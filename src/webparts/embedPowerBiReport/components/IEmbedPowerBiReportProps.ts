import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { ServiceScope } from "@microsoft/sp-core-library";

export interface IEmbedPowerBiReportProps {
  webPartContext: IWebPartContext;
  serviceScope: ServiceScope;
  defaultWorkspaceId: string;
  defaultReportId: string;
  defaultWidthToHeight: number;
}

export interface IEmbedPowerBiReportState {
  loading: boolean;
  workspaceId: string;
  reportId: string;
  widthToHeight: number;
}
