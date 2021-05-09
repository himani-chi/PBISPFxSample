import * as React from 'react';
import styles from './EmbedPowerBiReport.module.scss';
import { IEmbedPowerBiReportProps, IEmbedPowerBiReportState } from './IEmbedPowerBiReportProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { PowerBiWorkspace, PowerBiReport } from './../../../models/PowerBiModels';
import { PowerBiService } from './../../../services/PowerBiService';
import { PowerBiEmbeddingService } from './../../../services/PowerBiEmbeddingService';

export default class EmbedPowerBiReport extends React.Component<IEmbedPowerBiReportProps, IEmbedPowerBiReportState> {

  constructor(props: IEmbedPowerBiReportProps){
    super(props);
  }

  public state: IEmbedPowerBiReportState = {
    workspaceId: this.props.defaultWorkspaceId,
    reportId: this.props.defaultReportId,
    widthToHeight: this.props.defaultWidthToHeight,
    loading: false
  };

  private reportCannotRender(): Boolean {
    return ((this.state.workspaceId === undefined) || (this.state.workspaceId === "")) ||
      ((this.state.reportId === undefined) || (this.state.reportId === ""));
  }

  public render(): React.ReactElement<IEmbedPowerBiReportProps> {

    let containerHeight = this.props.webPartContext.domElement.clientWidth / (this.state.widthToHeight/100);

    console.log("EmbedPowerBiReport.render");

    return (
      <div className={styles.embedPowerBiReport}  >
      {this.state.loading ? (
        <div id="loading" className={styles.loadingContainer} >Calling to Power BI Service</div> 
      ) : ( 
        this.reportCannotRender() ? 
        <div id="message-container" className={styles.messageContainer} >Select a workspace and report from the web part property pane</div> : 
        <div id="embed-container" className={styles.embedContainer} style={{height: containerHeight }} ></div> 
      )}        
    </div>
    );
  }

  public componentDidMount() {
    console.log("componentDidUpdate");
    this.embedReport();
  }

  public componentDidUpdate(prevProps: IEmbedPowerBiReportProps, prevState: IEmbedPowerBiReportState, prevContext: any): void {
    console.log("componentDidUpdate");
    this.embedReport();
  }

  private embedReport() {
    let embedTarget: HTMLElement = document.getElementById('embed-container');
    if (!this.state.loading && !this.reportCannotRender()) {
      PowerBiService.GetReport(this.props.serviceScope, this.state.workspaceId, this.state.reportId).then((report: PowerBiReport) => {
        PowerBiEmbeddingService.embedReport(report, embedTarget);
      });
    }
  }

}
