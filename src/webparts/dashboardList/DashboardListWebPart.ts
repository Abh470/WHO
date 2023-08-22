import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

// import styles from './DashboardListWebPart.module.scss';
import * as strings from 'DashboardListWebPartStrings';
import * as $ from 'jquery';
//import * as moment from 'moment';
require("../../webparts/WHO-Imprest-Checklist/css/style.css");
require("../../webparts/WHO-Imprest-Checklist/css/padding.css");
require("../../webparts/WHO-Imprest-Checklist/js/common.js");
const WHOLogo: any = require('../../webparts/WHO-Imprest-Checklist/images/who-logo.png');
const ActionLogo: any = require('../../webparts/WHO-Imprest-Checklist/images/Action.png');
import { sp } from "@pnp/sp/presets/all";
import { SPComponentLoader } from '@microsoft/sp-loader';
import { AutoScroll } from 'office-ui-fabric-react';



var BlobUrlObject = [];

export interface IDashboardListWebPartProps {
  description: string;
}

export default class DashboardListWebPart extends BaseClientSideWebPart<IDashboardListWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
    sp.setup(this.context);
    return super.onInit();
  }



  public render(): void {
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css");
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css');
    SPComponentLoader.loadCss('https://code.jquery.com/jquery-1.12.4.min.js');
    // SPComponentLoader.loadCss("https://stackpath.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css");
    SPComponentLoader.loadCss("https://fonts.googleapis.com/css?family=Poppins:400,500,600,700&display=swap");
    SPComponentLoader.loadCss("https://cdn.datatables.net/1.13.1/css/jquery.dataTables.min.css");
    SPComponentLoader.loadScript("https://cdn.datatables.net/1.13.1/js/jquery.dataTables.min.js");
    this.domElement.innerHTML = `
    <nav class="navbar navbar-custom header-nav">
  <div class="container-fluid">
    <div class="navbar-header navbar-header-custom"> <a class="navbar-brand col-sm-3 col-xs-12" href="Titan-External-Employee-Portal.html"><img src="${WHOLogo}" class="logo" alt=""></a>
      <h2 class="m-0 col-sm-12 col-xs-12 text-center pt10 pb10">Imprest Checklist </h2>
      <div class="col-md-4 col-sm-12 col-xs-12 pr0 tpr15 tmb10">
        <div class="col-xs-12 external-emp-profile-box"> <a href="#"> <img src="" id="current-user-image" class="img-circle" alt=""> </a>
          <div class="ml15 mr15">
            <h3 id="text"></h3>
          </div>
        </div>
      </div> 
    </div>
  </div>
</nav>
<div class="container-fluid">
  <section class="preview mb-50 panel">
    <div class="message_panel">
      <div class="field-section row">
        <div class="col-sm-12">
          <div class="dashboard-top-section">
              <div class="new-button-section">
                <button class="btn btn-info" id="new-form">+ New</button>
              </div>
              ${ /* <div class="search-filter-section">
                <input class="form-control" type="text" placeholder="Search">
                <button class="btn btn-info ml10"><i class="fa fa-filter"></i> Filter</button>
              </div>*/''}
          </div>
          <div class="clearfix"></div>
          <div class="table-responsive mt-10">
            <table class="table table-bordered table-white" id="tableId">
              <thead class="blue_bg">
                <tr>
                  <th>Country</th>
                  <th>Imprest Account</th>
                  <th>GL Accounts</th>
                  <th>Month & Year</th>
                  <th>Imprest Checklist</th>
                  <th>Action</th>
                </tr>
              </thead>
              <tbody id="table-body">
            
              </tbody>
            </table>
            ${ /* <div class="col-sm-8 col-xs-12 pl0">
                <ul class="pagination custom-pagination font-14">
                  <li><a href="#"><i class="fa fa-angle-double-left"></i></a>
                  </li>
                  <li><a href="#"><i class="fa fa-angle-left"></i></a>
                  </li>
                  <li class="active"><a href="#">1</a></li>
                  <li><a href="#">2</a></li>
                  <li><a href="#">3</a></li>
                  <li><a href="#">4</a></li>
                  <li><a href="#"><i class="fa fa-angle-right"></i></a>
                  </li>
                  <li><a href="#"><i class="fa fa-angle-double-right"></i></a>
                  </li>
                </ul>
              </div>
  */''}
          </div>
          <div class="clearfix"></div>
        </div>
      </div>
    </div>
  </section>
</div>
    `;
    this._bindEvent();
  }


  private async _bindEvent() {
    this.wideSitePages();
    this.getCurrentUser();
    this.getImprestCheckListData();
    this.domElement.querySelector('#new-form').addEventListener('click', () => {
      window.location.href = `${this.context.pageContext.web.absoluteUrl}/SitePages/ApproverEditForm.aspx`;
    })

  }


  private wideSitePages() {
    $("#s4-workspace").hide();
    $("#spCommandBar").hide();
    $(".webPartContainer").hide();
    $("#CommentsWrapper").hide();
    $(".SideTabMenu").hide();
    $("#masterFooter").hide();
    $("#SuiteNavWrapper").hide();
    $("#sp-appBar").hide();
    $("#spLeftNav").hide();
    $("#spTopPlaceholder").hide();
    $("#spSiteHeader").hide();
    $('#spPageCanvasContent').find('[data-automation-id="CanvasZone"]>div').addClass("sp-custom-main-box");
  }


  private async getImprestCheckListData() {
    let items = await sp.web.lists.getByTitle("ImprestCheckList").items.orderBy("ID", true).getAll();
    console.log(items)
    let html = ``;
    items.forEach((item) => {
      let date = new Date(item.Month_x0020_and_x0020_Year);
      const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
      const month = monthNames[date.getMonth()]; // get month name
      const year = date.getFullYear(); // get year
      const formattedDate = month + "-" + year; // concatenate month name and year
      console.log(formattedDate);
      // let GlAccount = item.GL_x0020_account_x0028_s_x0029_;
      let GlAccount1 = item.GLaccount_x0028_s_x0029_;
      html += `
      <tr>
          <td>${item.Country}</td>
          <td>${item.Imprest_x0020_Account}</td>
          <td>${GlAccount1}</td>
          <td>${formattedDate}</td>
          <td>${item.ImprestChecklistStatus}</td>
          <td style="font-size: 18px;">
          <a href='${this.context.pageContext.web.absoluteUrl}/SitePages/ApproverEditForm.aspx?Imprestid=${item.Id}'><img src="${ActionLogo}" width="16" height="16" alt=""/></a>
          </td>
    </tr>
      `;
    })
    await $("#table-body").html(html);
    ($("#tableId") as any).DataTable({
      items: 100,
      itemsOnPage: 10,
      cssStyle: 'light-theme',
      scrollY: '500px',
      scrollX: true,
      sScrollXInner: "100%",
      //bFilter: false
    });
    // var table = ($('#tableId')as any).DataTable();
    // table.columns.adjust();
    // $('.dataTable').wrap('<div class="dataTables_scroll" />');

    //   $('#example').DataTable( {
    //     responsive: true
    // } );
  }

  private async getCurrentUser() {
    //let user = await sp.web.currentUser();
    $("#text").text(this.context.pageContext.user.displayName);
    var imggurl = await this.blob(this.context.pageContext.user.email);
    $("#current-user-image").attr("src", imggurl);
  }



  private async blob(userPrincipalName): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      var CacheblobUrl;
      function userExists(userPrincipalName) {
        return BlobUrlObject.some(function (el, index) {
          if (el.Email === userPrincipalName) {
            CacheblobUrl = el.BlobUrl
          }

          return el.Email === userPrincipalName;
        });
      }
      //  let userPrincipalName = "Abhishek.Negi@titan4work.com";
      if (userExists(userPrincipalName) == false) {
        this.context.msGraphClientFactory
          .getClient()
          .then((client): void => {

            client.api(`users/${userPrincipalName}/photo/$value`)
              .responseType('blob').version("beta").get().then((response: Blob) => {
                const blobUrl = window.URL.createObjectURL(response);
                var img = document.createElement('img');
                img.src = blobUrl
                img.className = "img-responsive";
                BlobUrlObject.push({ BlobUrl: img.src, Email: userPrincipalName })
                resolve(img.src)

              }).catch((error) => {
                console.log(error)
                let errorurl = "https://t4.ftcdn.net/jpg/00/64/67/63/360_F_64676383_LdbmhiNM6Ypzb3FM4PPuFP9rHe7ri8Ju.jpg";
                BlobUrlObject.push({ BlobUrl: errorurl, Email: userPrincipalName });
                resolve('https://t4.ftcdn.net/jpg/00/64/67/63/360_F_64676383_LdbmhiNM6Ypzb3FM4PPuFP9rHe7ri8Ju.jpg')
              })

          })
      }
      else {
        resolve(CacheblobUrl)
        console.log("Blob Url already")
      }
    })

  }
  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
