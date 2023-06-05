import * as React from 'react';
// import styles from './ApproverEditForm.module.scss';
import { IApproverEditFormProps } from './IApproverEditFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp/presets/all";
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from 'jquery';
require("../../WHO-Imprest-Checklist/css/style.css");
require("../../WHO-Imprest-Checklist/css/padding.css");
require("../../WHO-Imprest-Checklist/js/common.js");
require("../../WHO-Imprest-Checklist/src/richtext.min.css");
require("../../WHO-Imprest-Checklist/src/jquery.richtext.js")
require("../../WHO-Imprest-Checklist/css/site.css");
// import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
// import jsPDF from 'jspdf';
// import html2canvas from 'html2canvas';

const WHOLogo: any = require('../../WHO-Imprest-Checklist/images/who-logo.png');
// import "../../WHO-Imprest-Checklist/css/owl.carousel.min.css";


var BlobUrlObject = [];

//For attachments
var fileInfos = [];

//Imprest id by query string
var itemId;

export interface IDropdown {
  ID: any,
  CheckList: any
}

export interface State {
  Imgsrc: "",
  richtextdescription1: string,
  richtextdescription2: string,
  richtextdescription3: string,
  richtextdescription4: string,
  richtextdescription5: string,
  richtextdescription6: string,
  richtextdescription7: string,
  richtextdescription8: string,
  richtextdescription9: string,
  richtextdescription10: string,
  richtextdescription11: string,
  ImprestCheck1: string,
  ImprestCheck2: string,
  ImprestCheck3: string,
  ImprestCheck4: string,
  ImprestCheck5: string,
  ImprestCheck6: string,
  ImprestCheck7: string,
  ImprestCheck8: string,
  ImprestCheck9: string,
  ImprestCheck10: string,
  ImprestCheck11: string,
  country: string,
  monthandYear: any,
  imprestAccount: string,
  GLAccount: string,
  Feedback: string,
  FileArr: any[],
  Disabled: boolean,
  CheckListDropdown: IDropdown[]
  UpdateEditForm: boolean

}
const initialstate = {
  // Imgsrc: "",
  richtextdescription1: "",
  richtextdescription2: "",
  richtextdescription3: "",
  richtextdescription4: "",
  richtextdescription5: "",
  richtextdescription6: "",
  richtextdescription7: "",
  richtextdescription8: "",
  richtextdescription9: "",
  richtextdescription10: "",
  richtextdescription11: "",
  country: "",
  monthandYear: "",
  imprestAccount: "",
  GLAccount: "",
  Feedback: "",
  FileArr: [],
  Disabled: false,
  // CheckListDropdown: [],
  ImprestCheck1: null,
  ImprestCheck2: null,
  ImprestCheck3: null,
  ImprestCheck4: null,
  ImprestCheck5: null,
  ImprestCheck6: null,
  ImprestCheck7: null,
  ImprestCheck8: null,
  ImprestCheck9: null,
  ImprestCheck10: null,
  ImprestCheck11: null,
  UpdateEditForm: false

};

export default class ApproverEditForm extends React.Component<IApproverEditFormProps, State> {

  constructor(props: IApproverEditFormProps, state: State) {
    super(props);
    sp.setup(props.context);
    this.state = {
      Imgsrc: "",
      richtextdescription1: "",
      richtextdescription2: "",
      richtextdescription3: "",
      richtextdescription4: "",
      richtextdescription5: "",
      richtextdescription6: "",
      richtextdescription7: "",
      richtextdescription8: "",
      richtextdescription9: "",
      richtextdescription10: "",
      richtextdescription11: "",
      country: "",
      monthandYear: "",
      imprestAccount: "",
      GLAccount: "",
      Feedback: "",
      FileArr: [],
      Disabled: false,
      CheckListDropdown: [],
      ImprestCheck1: null,
      ImprestCheck2: null,
      ImprestCheck3: null,
      ImprestCheck4: null,
      ImprestCheck5: null,
      ImprestCheck6: null,
      ImprestCheck7: null,
      ImprestCheck8: null,
      ImprestCheck9: null,
      ImprestCheck10: null,
      ImprestCheck11: null,
      UpdateEditForm: false

    };
  }


  public componentDidMount(): void {
    $("#loader").hide();
    this.getCurrentUser();
    itemId = this.getParameterByName("Imprestid");
    if (itemId) {
      this.getCheckListItemByID(itemId);
      this.setState({ UpdateEditForm: true })
    }
    this.getCheckListDropdown();
    setTimeout(() => {
      ($('.content1') as any).richText();
      ($('.content2') as any).richText();
      ($('.content3') as any).richText();
      ($('.content4') as any).richText();
      ($('.content5') as any).richText();
      ($('.content6') as any).richText();
      ($('.content7') as any).richText();
      ($('.content8') as any).richText();
      ($('.content9') as any).richText();
      ($('.content10') as any).richText();
      ($('.content11') as any).richText();
      if (itemId) {
        $("textarea.content1").siblings("div.richText-editor").html(this.state.richtextdescription1);
        $("textarea.content2").siblings("div.richText-editor").html(this.state.richtextdescription2);
        $("textarea.content3").siblings("div.richText-editor").html(this.state.richtextdescription3);
        $("textarea.content4").siblings("div.richText-editor").html(this.state.richtextdescription4);
        $("textarea.content5").siblings("div.richText-editor").html(this.state.richtextdescription5);
        $("textarea.content6").siblings("div.richText-editor").html(this.state.richtextdescription6);
        $("textarea.content7").siblings("div.richText-editor").html(this.state.richtextdescription7);
        $("textarea.content8").siblings("div.richText-editor").html(this.state.richtextdescription8);
        $("textarea.content9").siblings("div.richText-editor").html(this.state.richtextdescription9);
        $("textarea.content10").siblings("div.richText-editor").html(this.state.richtextdescription10);
        $("textarea.content11").siblings("div.richText-editor").html(this.state.richtextdescription11);

        $("textarea.content1").siblings("div.richText-editor").attr("contenteditable", (!this.state.Disabled as any));
        $("textarea.content2").siblings("div.richText-editor").attr("contenteditable", (!this.state.Disabled as any));
        $("textarea.content3").siblings("div.richText-editor").attr("contenteditable", (!this.state.Disabled as any));
        $("textarea.content4").siblings("div.richText-editor").attr("contenteditable", (!this.state.Disabled as any));
        $("textarea.content5").siblings("div.richText-editor").attr("contenteditable", (!this.state.Disabled as any));
        $("textarea.content6").siblings("div.richText-editor").attr("contenteditable", (!this.state.Disabled as any));
        $("textarea.content7").siblings("div.richText-editor").attr("contenteditable", (!this.state.Disabled as any));
        $("textarea.content8").siblings("div.richText-editor").attr("contenteditable", (!this.state.Disabled as any));
        $("textarea.content9").siblings("div.richText-editor").attr("contenteditable", (!this.state.Disabled as any));
        $("textarea.content10").siblings("div.richText-editor").attr("contenteditable", (!this.state.Disabled as any));
        $("textarea.content11").siblings("div.richText-editor").attr("contenteditable", (!this.state.Disabled as any));
      }

    }, 2000);

  }


  private printpdf = (e) => {
    $("#loader").show();
    $.ajax({
      type: "POST",
      crossDomain: true,
      headers: {
        'Access-Control-Allow-Origin': '*',
      },
      contentType: "application/json",
      //dataType: 'native',
      url: "https://prod-05.centralindia.logic.azure.com:443/workflows/4fe4a4a4715042b08044b9890c8a272c/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=v-d6nNrZfwkSw-gDhYBJ_UHu3g_JlfLy4PebA-uf3-U",
      xhrFields: {
        responseType: 'blob'
      },
      data: JSON.stringify({ "id": Number(itemId) })
      ,
      success: function (blob) {
        console.log(blob.size);
        var link = document.createElement('a');
        link.href = window.URL.createObjectURL(blob);
        link.download = "Imprest Checklist Form" + new Date() + ".pdf";
        link.click();
        $("#loader").hide();
      },
      error: function (error) {
        console.log(JSON.stringify(error));
        $("#loader").hide();
      }
    });

    // var printContents ='';
    // printContents += $(".header-nav").html();
    //  printContents += document.getElementById("mypdf").innerHTML;
    // 	var originalContents = document.body.innerHTML;

    // 	document.body.innerHTML = printContents;

    // 	 window.print();
    //    window.location.reload();
    // document.body.innerHTML = originalContents;
  }

  // public documentprint = (e) => {
  //   e.preventDefault();
  //   window.print()
  //   const myinput = document.getElementById('mypdf');
  //   //window.scrollTo(0,0); 

  //   var HTML_Width = $("#mypdf").width();
  //   var HTML_Height = $("#mypdf").height();
  //   html2canvas(myinput, {
  //     allowTaint: true,
  //     useCORS: true,
  //     windowWidth: HTML_Width,
  //     windowHeight: HTML_Height,
  //     // scrollY: -window.scrollY,
  //   })
  //     .then((canvas) => {
  //       // window.scrollTo(0, document.body.scrollHeight || document.documentElement.scrollHeight);
  //       var imgWidth = 200;
  //       var pageHeight = 290;
  //       var imgHeight = canvas.height * imgWidth / canvas.width;
  //       var heightLeft = imgHeight;
  //       const imgData = canvas.toDataURL('image/png');
  //       const mynewpdf = new jsPDF('p', 'mm', 'a4');
  //       var position = 0;
  //       var top_left_margin = 15;
  //       var PDF_Width = HTML_Width + (top_left_margin * 2);
  //       var PDF_Height = (PDF_Width * 1.5) + (top_left_margin * 2);
  //       var totalPDFPages = Math.ceil(HTML_Height / PDF_Height) - 1;
  //       //mynewpdf.addImage(imgData, 'JPEG', 5, position, imgWidth, imgHeight);
  //       for (var i = 1; i <= totalPDFPages; i++) {
  //         mynewpdf.addPage(PDF_Width.toString(), PDF_Height.toString() as any);
  //         mynewpdf.addImage(imgData, 'JPEG', 5, position, imgWidth, imgHeight);
  //         // mynewpdf.addImage(imgData, 'JPG', top_left_margin, -(PDF_Height*i)+(top_left_margin*4),HTML_Width,HTML_Height);
  //       }
  //       mynewpdf.save("ApproverEditForm.pdf");
  //     });


  //   // var HTML_Width = $("#mypdf").width();
  //   // var HTML_Height = $("#mypdf").height();
  //   // var top_left_margin = 15;
  //   // var PDF_Width = HTML_Width + (top_left_margin * 2);
  //   // var PDF_Height = (PDF_Width * 1.5) + (top_left_margin * 2);
  //   // var canvas_image_width = HTML_Width;
  //   // var canvas_image_height = HTML_Height;

  //   // var totalPDFPages = Math.ceil(HTML_Height / PDF_Height) - 1;


  //   // html2canvas($("#mypdf")[0], { allowTaint: true, useCORS: true, }).then(function (canvas) {
  //   //   canvas.getContext('2d');

  //   //   console.log(canvas.height + "  " + canvas.width);


  //   //   var imgData = canvas.toDataURL("image/jpeg", 1.0);
  //   //   var pdf = new jsPDF('p', 'pt', [PDF_Width, PDF_Height]);
  //   //   pdf.addImage(imgData, 'JPG', top_left_margin, top_left_margin, canvas_image_width, canvas_image_height);


  //   //   for (var i = 1; i <= totalPDFPages; i++) {
  //   //     pdf.addPage(PDF_Width.toString(),PDF_Height.toString() as any);
  //   //     pdf.addImage(imgData, 'JPG', top_left_margin, top_left_margin, canvas_image_width, canvas_image_height);
  //   //     //pdf.addImage(imgData, 'JPG', top_left_margin, -(PDF_Height*i)+(top_left_margin*4),canvas_image_width,canvas_image_height);
  //   //   }

  //   //   pdf.save("HTML-Document.pdf");
  //   // });

  // }


  private async getCheckListItemByID(id) {
    let items = await sp.web.lists.getByTitle("ImprestCheckList").items.getById(id).get();
    console.log(items)
    let files = await sp.web.lists.getByTitle("ImprestListFiles").items.filter(`ImprestCheckList eq ${items.ID}`).select("FileLeafRef,Editor/Name").expand("Editor").get();
    console.log(files);
    let UpdatedFileInfos = [];
    files.forEach(async (file) => {
      const user = await sp.web.ensureUser(file.Editor.Name)
      fileInfos.push({
        name: file.FileLeafRef,
        content: "",
        person: user.data.Title
      });
      this.setState({ FileArr: fileInfos })
    })

    this.setState({ country: items.Country });
    this.setState({ imprestAccount: items.Imprest_x0020_Account });
    this.setState({ GLAccount: items.GLaccount_x0028_s_x0029_ });
    this.setState({ ImprestCheck1: items.Imprest_x0020_Check1Id });
    this.setState({ ImprestCheck2: items.Imprest_x0020_Check2Id });
    this.setState({ ImprestCheck3: items.Imprest_x0020_Check3Id });
    this.setState({ ImprestCheck4: items.Imprest_x0020_Check4Id });
    this.setState({ ImprestCheck5: items.Imprest_x0020_Check5Id });
    this.setState({ ImprestCheck6: items.Imprest_x0020_Check6Id });
    this.setState({ ImprestCheck7: items.Imprest_x0020_Check7Id });
    this.setState({ ImprestCheck8: items.Imprest_x0020_Check8Id });
    this.setState({ ImprestCheck9: items.Imprest_x0020_Check9Id });
    this.setState({ ImprestCheck10: items.Imprest_x0020_Check10Id });
    this.setState({ ImprestCheck11: items.Imprest_x0020_Check11Id });
    this.setState({ richtextdescription1: items.Imprest_x0020_Remarks1 });
    this.setState({ richtextdescription2: items.Imprest_x0020_Remarks2 });
    this.setState({ richtextdescription3: items.Imprest_x0020_Remarks3 });
    this.setState({ richtextdescription4: items.Imprest_x0020_Remarks4 });
    this.setState({ richtextdescription5: items.Imprest_x0020_Remarks5 });
    this.setState({ richtextdescription6: items.Imprest_x0020_Remarks6 });
    this.setState({ richtextdescription7: items.Imprest_x0020_Remarks7 });
    this.setState({ richtextdescription8: items.Imprest_x0020_Remarks8 });
    this.setState({ richtextdescription9: items.Imprest_x0020_Remarks9 });
    this.setState({ richtextdescription10: items.Imprest_x0020_Remarks10 });
    this.setState({ richtextdescription11: items.Imprest_x0020_Remarks11 });



    this.setState({ Feedback: items.ImprestFeedback });

    let date = items.Month_x0020_and_x0020_Year;
    let dt = date.split("T")[0]

    //let dt = date.getFullYear() + "-" + date.getMonth() + "-" + date.getDate()
    this.setState({ monthandYear: dt });
    // this.setState({});
    if (items.ImprestChecklistStatus == "Draft") {
      this.setState({ Disabled: false });
    }
    else if (items.ImprestChecklistStatus == "Completed") {
      this.setState({ Disabled: true });
    }
  }

  private async getCurrentUser() {
    //let user = await sp.web.currentUser();
    $("#text").text(this.props.context.pageContext.user.displayName);
    var imggurl = await this.blob(this.props.context.pageContext.user.email);
    //$("#current-user-image").attr("src", imggurl);
    this.setState({ Imgsrc: imggurl })


    $("textarea").each(function () {
      this.setAttribute("style", "height:" + (this.scrollHeight) + "px;overflow-y:hidden; resize:none; min-height:75px;");
    }).on("input", function () {
      this.style.height = "0";
      this.style.height = (this.scrollHeight) + "px";
    });
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
        this.props.context.msGraphClientFactory
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
  private onTextChange = (newText: string) => {
    return newText;
  }

  private CheckValidation(): Promise<boolean> {
    return new Promise<boolean>((resolve, reject) => {
      if (this.state.country == "") {
        alert("Please Fill the Country Name.");
        resolve(false);
      }
      else if (this.state.monthandYear == "") {
        alert("Please Fill the Month & Year.");
        resolve(false);
      }
      else if (this.state.imprestAccount == "") {
        alert("Please Fill Imprest Account.");
        resolve(false);
      }
      else if (this.state.GLAccount == "") {
        alert("Please Fill the GL Account.");
        resolve(false);
      }
      else {
        resolve(true)
      }

    })

  }

  private submitTask(checkListStatus, completedstatus): Promise<void> {
    return new Promise<void>(async (resolve, reject) => {
      let today = new Date();
      let validate = await this.CheckValidation();
      if (validate) {
        $("#loader").show();
        sp.web.lists.getByTitle('ImprestCheckList').items.add({
          Country: this.state.country,
          Imprest_x0020_Account: this.state.imprestAccount,
          GLaccount_x0028_s_x0029_: this.state.GLAccount,
          Month_x0020_and_x0020_Year: this.state.monthandYear,
          CompletedStatus: completedstatus,
          ImprestChecklistStatus: checkListStatus,
          // Imprest_x0020_Remarks1: this.state.richtextdescription1,
          // Imprest_x0020_Remarks2: this.state.richtextdescription2,
          // Imprest_x0020_Remarks3: this.state.richtextdescription3,
          // Imprest_x0020_Remarks4: this.state.richtextdescription4,
          // Imprest_x0020_Remarks5: this.state.richtextdescription5,
          // Imprest_x0020_Remarks6: this.state.richtextdescription6,
          // Imprest_x0020_Remarks7: this.state.richtextdescription7,
          // Imprest_x0020_Remarks8: this.state.richtextdescription8,
          // Imprest_x0020_Remarks9: this.state.richtextdescription9,
          // Imprest_x0020_Remarks10: this.state.richtextdescription10,
          // Imprest_x0020_Remarks11: this.state.richtextdescription11,
          Imprest_x0020_Remarks1: $(".content1").val(),
          Imprest_x0020_Check1Id: this.state.ImprestCheck1,
          Imprest_x0020_Check2Id: this.state.ImprestCheck2,
          Imprest_x0020_Check3Id: this.state.ImprestCheck3,
          Imprest_x0020_Check4Id: this.state.ImprestCheck4,
          Imprest_x0020_Check5Id: this.state.ImprestCheck5,
          Imprest_x0020_Check6Id: this.state.ImprestCheck6,
          Imprest_x0020_Check7Id: this.state.ImprestCheck7,
          Imprest_x0020_Check8Id: this.state.ImprestCheck8,
          Imprest_x0020_Check9Id: this.state.ImprestCheck9,
          Imprest_x0020_Check10Id: this.state.ImprestCheck10,
          Imprest_x0020_Check11Id: this.state.ImprestCheck11,
          ImprestFeedback: this.state.Feedback,
          SubmittedDate: today

        }).then(async (response) => {
          console.log(response);
          this.state.FileArr.forEach(async (file) => {
            sp.web.getFolderByServerRelativeUrl("ImprestListFiles").files.add(file.name, file.content, true)
              .then(async (result) => {
                await result.file.getItem().then(item => {
                  item.update({
                    ImprestCheckListId: response.data.Id
                  })
                  .catch((err) => alert(JSON.stringify(err)))
                })
              })
          })
          await this.UpdateRichtextPickerFieldInLoop(response.data.Id);
         // this.setState(initialstate);
        })
        .then(()=>{
          setTimeout(() => {
            $("#loader").hide();
            alert("Form has been submitted successfully");
            window.location.href =`${this.props.context.pageContext.web.absoluteUrl}/SitePages/DashboardList.aspx`;         
          }, this.state.FileArr.length+1 * 1000);
        })
        .catch((err) => console.log(err));
      }

    })
  }

  private async UpdateRichtextPickerFieldInLoop(id) {
    if ($(".content2").val() != '<div><br></div>' || $(".content3").val() != '<div><br></div>') {
     await sp.web.lists.getByTitle('ImprestCheckList').items.getById(id).update({
        Imprest_x0020_Remarks2: $(".content2").val(),
        Imprest_x0020_Remarks3: $(".content3").val()
      })
    }
    if ($(".content4").val() != '<div><br></div>' || $(".content5").val() != '<div><br></div>') {
      await sp.web.lists.getByTitle('ImprestCheckList').items.getById(id).update({
        Imprest_x0020_Remarks4: $(".content4").val(),
        Imprest_x0020_Remarks5: $(".content5").val(),
      })
    }
    if ($(".content6").val() != '<div><br></div>'|| $(".content7").val() != '<div><br></div>') {
     await sp.web.lists.getByTitle('ImprestCheckList').items.getById(id).update({
        Imprest_x0020_Remarks6: $(".content6").val(),
        Imprest_x0020_Remarks7: $(".content7").val(),
      })
    }
    if ($(".content8").val() != '<div><br></div>'|| $(".content9").val() != '<div><br></div>') {
      await sp.web.lists.getByTitle('ImprestCheckList').items.getById(id).update({
         Imprest_x0020_Remarks8: $(".content8").val(),
         Imprest_x0020_Remarks9: $(".content9").val(),
       })
     }
    if ($(".content10").val() != '<div><br></div>' ||$(".content11").val() != '<div><br></div>') {
      await sp.web.lists.getByTitle('ImprestCheckList').items.getById(id).update({
        Imprest_x0020_Remarks10: $(".content10").val(),
        Imprest_x0020_Remarks11: $(".content11").val(),
      })
    }

  }

  private async deleteFile(id, fileName) {
    sp.web.getFileByServerRelativeUrl(`${this.props.context.pageContext.web.serverRelativeUrl}/ImprestListFiles/${fileName}`).
      recycle().then(function (data) {
        console.log(data + "file deleted");
      });
    // const item = await sp.web.lists.getByTitle("ImprestListFiles").items.getById(id).get();
    // await item.attachmentFiles.getByName(fileName).recycle();
    // console.log(fileName + "file Deleted")

  }
  private updateTask(itemId, checkListStatus, completedstatus): Promise<void> {
    return new Promise<void>(async (resolve, reject) => {
      let today = new Date();
      let validate = await this.CheckValidation();
      if (validate) {
        $('#loader').show();
        sp.web.lists.getByTitle('ImprestCheckList').items.getById(itemId).update({
          Country: this.state.country,
          Imprest_x0020_Account: this.state.imprestAccount,
          GLaccount_x0028_s_x0029_: this.state.GLAccount,
          Month_x0020_and_x0020_Year: this.state.monthandYear,
          CompletedStatus: completedstatus,
          ImprestChecklistStatus: checkListStatus,
          // Imprest_x0020_Remarks1: this.state.richtextdescription1,
          // Imprest_x0020_Remarks2: this.state.richtextdescription2,
          // Imprest_x0020_Remarks3: this.state.richtextdescription3,
          // Imprest_x0020_Remarks4: this.state.richtextdescription4,
          // Imprest_x0020_Remarks5: this.state.richtextdescription5,
          // Imprest_x0020_Remarks6: this.state.richtextdescription6,
          // Imprest_x0020_Remarks7: this.state.richtextdescription7,
          // Imprest_x0020_Remarks8: this.state.richtextdescription8,
          // Imprest_x0020_Remarks9: this.state.richtextdescription9,
          // Imprest_x0020_Remarks10: this.state.richtextdescription10,
          // Imprest_x0020_Remarks11: this.state.richtextdescription11,
          Imprest_x0020_Remarks1: $(".content1").val(),

          Imprest_x0020_Check1Id: this.state.ImprestCheck1,
          Imprest_x0020_Check2Id: this.state.ImprestCheck2,
          Imprest_x0020_Check3Id: this.state.ImprestCheck3,
          Imprest_x0020_Check4Id: this.state.ImprestCheck4,
          Imprest_x0020_Check5Id: this.state.ImprestCheck5,
          Imprest_x0020_Check6Id: this.state.ImprestCheck6,
          Imprest_x0020_Check7Id: this.state.ImprestCheck7,
          Imprest_x0020_Check8Id: this.state.ImprestCheck8,
          Imprest_x0020_Check9Id: this.state.ImprestCheck9,
          Imprest_x0020_Check10Id: this.state.ImprestCheck10,
          Imprest_x0020_Check11Id: this.state.ImprestCheck11,
          ImprestFeedback: this.state.Feedback,
          SubmittedDate: today

        }).then(async (response) => {
          console.log(response);
          this.state.FileArr.forEach(async (file) => {
            if (file.content != "") {
              sp.web.getFolderByServerRelativeUrl("ImprestListFiles").files.add(file.name, file.content, true)
                .then(async (result) => {
                  await result.file.getItem().then(item => {
                    item.update({
                      ImprestCheckListId: itemId
                    }).catch((error) => console.log(error))
                  })
                })
            }
          })
          await this.UpdateRichtextPickerFieldInLoop(itemId);
          //this.setState(initialstate);
        })
        .then(()=>{
          setTimeout(()=>{
            $("#loader").hide();
            alert("Form has been saved successfully.");
            window.location.href =`${this.props.context.pageContext.web.absoluteUrl}/SitePages/DashboardList.aspx`;
          },this.state.FileArr.length +1 *1000)
        })
          .catch((err) => alert(JSON.stringify(err)));
      }
    })
  }

  private async getCheckListDropdown() {
    let items: [] = await sp.web.lists.getByTitle("CheckList").items.get();
    this.setState({ CheckListDropdown: items })
    //console.log(this.state.CheckListDropdown);

  }

  private setButtonsEventHandlersForAttachment(files): void {
    if (files) {
      var fileCount = files.length;
      console.log(fileCount);
      let scope = this;
      for (var i = 0; i < fileCount; i++) {
        //var fileName = files[i].name;
        //console.log(fileName);
        var file = files[i];
        var reader = new FileReader();
        reader.onload = (function (file) {
          return function (e) {
            //console.log(file.name);
            //Push the converted file into array
            fileInfos.push({
              name: file.name,
              content: e.target.result
            });
            console.log(fileInfos);
            scope.setState({ FileArr: fileInfos });
          }
        })(file);
        reader.readAsArrayBuffer(file);

      }
    }
  }

  //Query parameter
  private getParameterByName(name, url = window.location.href) {
    name = name.replace(/[\[\]]/g, '\\$&');
    var regex = new RegExp('[?&]' + name + '(=([^&#]*)|&|#|$)'),
      results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, ' '));
  }

  public render(): React.ReactElement<IApproverEditFormProps> {
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css");
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css');
    SPComponentLoader.loadCss('https://code.jquery.com/jquery-1.12.4.min.js');

    return (
      <>
        <div id="mypdf">
          <nav className="navbar navbar-custom header-nav" >
            <div className="container-fluid">
              <div className="navbar-header navbar-header-custom"> <a className="navbar-brand col-sm-3 col-xs-12" href="Titan-External-Employee-Portal.html">
                <img src={WHOLogo} className="logo" alt="" crossOrigin='anonymous' /></a>
                <h2 className="m-0 col-sm-12 col-xs-12 text-center pt10 pb10">Imprest Checklist </h2>
                <div className="col-md-4 col-sm-12 col-xs-12 pr0 tpr15 tmb10">
                  <div className="col-xs-12 external-emp-profile-box">
                    <img src={this.state.Imgsrc} id="current-user-image" className="img-circle" crossOrigin='anonymous' alt="" />
                    <div className="ml15 mr15">
                      <h3 id="text">Patrick Moore</h3>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </nav>
          <div className="container-fluid">
            <section className="preview mb-50 panel">
              <div className="message_panel imprest-panel">
                <div className="field-section row">
                  <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12">
                    <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12 mt-10 p-0">
                      <div className="row">
                        <div className="col-lg-6 col-md-6 col-sm-6 col-xs-12">
                          <div className="row">
                            <div className="form-group">
                              <label className="col-lg-3 col-md-4 col-sm-4 col-xs-12">Country <span style={{ color: "#e00" }}>*</span></label>
                              <div className="col-lg-9 col-md-8 col-sm-8 col-xs-12">
                                <input className="form-control" type="text" disabled={this.state.Disabled} value={this.state.country}
                                  onChange={(e) => this.setState({ country: e.target.value })} />
                              </div>
                            </div>
                          </div>
                        </div>
                        <div className="col-lg-6 col-md-6 col-sm-6 col-xs-12">
                          <div className="row">
                            <div className="form-group">
                              <label className="col-lg-3 col-md-4 col-sm-4 col-xs-12">Month & Year <span style={{ color: "#e00" }}>*</span></label>
                              <div className="col-lg-9 col-md-8 col-sm-8 col-xs-12">
                                <input className="form-control" type="date" disabled={this.state.Disabled}
                                  defaultValue={this.state.monthandYear}
                                  onChange={(e) => this.setState({ monthandYear: e.target.value })} />
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                      <div className="row mt-15">
                        <div className="col-lg-6 col-md-6 col-sm-6 col-xs-12">
                          <div className="row">
                            <div className="form-group">
                              <label className="col-lg-3 col-md-4 col-sm-4 col-xs-12">Imprest Account <span style={{ color: "#e00" }}>*</span></label>
                              <div className="col-lg-9 col-md-8 col-sm-8 col-xs-12">
                                <input className="form-control" type="text" disabled={this.state.Disabled} value={this.state.imprestAccount}
                                  onChange={(e) => this.setState({ imprestAccount: e.target.value })} />
                              </div>
                            </div>
                          </div>
                        </div>
                        <div className="col-lg-6 col-md-6 col-sm-6 col-xs-12">
                          <div className="row">
                            <div className="form-group">
                              <label className="col-lg-3 col-md-4 col-sm-4 col-xs-12">GL Account(s) <span style={{ color: "#e00" }}>*</span></label>
                              <div className="col-lg-9 col-md-8 col-sm-8 col-xs-12">
                                <input className="form-control" type="text" disabled={this.state.Disabled} value={this.state.GLAccount}
                                  onChange={(e) => this.setState({ GLAccount: e.target.value })} />
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                    <div className="clearfix"></div>
                    <div className="table-responsive mt-15">
                      <table className="table table-bordered table-white">
                        <thead className="blue_bg">
                          <tr>
                            <th className="tw-5">S No.</th>
                            <th>Checklist Item</th>
                            <th className="tw-15">Status</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <td>1</td>
                            <td>
                              <p>Is each transaction supported by appropriate supporting documents?</p>
                              {/* <RichText value={this.state.richtextdescription1}
                                isEditMode={!this.state.Disabled}
                                onChange={(text) => {
                                  this.setState({ richtextdescription1: text });
                                  return text;
                                }}
                              /> */}
                              <textarea className="content1" name="example1"
                                onChange={(e) => this.setState({ richtextdescription1: e.target.value })} />
                            </td>
                            <td>
                              <select className="form-control" disabled={this.state.Disabled}
                                value={this.state.ImprestCheck1} onChange={(e) => this.setState({ ImprestCheck1: e.target.value })}>
                                <option value={''}>--Select--</option>
                                {this.state.CheckListDropdown.map((items) => {
                                  return (
                                    <option value={items.ID}>{items.CheckList}</option>
                                  )
                                })}
                              </select>
                            </td>
                          </tr>
                          <tr>
                            <td>2</td>
                            <td>
                              <p>Payment per transaction charged to IPO does not exceed $2,500</p>
                              {/* <RichText value={this.state.richtextdescription2}
                                isEditMode={!this.state.Disabled}
                                onChange={(text) => {
                                  this.setState({ richtextdescription2: text });
                                  return text;
                                }}
                              /> */}
                              <textarea className="content2" name="example2"
                                onChange={(e) => this.setState({ richtextdescription2: e.target.value })} />
                            </td>
                            <td>
                              <select className="form-control" disabled={this.state.Disabled}
                                value={this.state.ImprestCheck2} onChange={(e) => this.setState({ ImprestCheck2: e.target.value })}>
                                <option value={''}>--Select--</option>
                                {this.state.CheckListDropdown.map((items) => {
                                  return (
                                    <option value={items.ID}>{items.CheckList}</option>
                                  )
                                })}
                              </select>
                            </td>
                          </tr>
                          <tr>
                            <td>3</td>
                            <td>
                              <p>Payments charged/receipts credited to IPOs are in line with Imprest PO guidelines.</p>
                              {/* <RichText value={this.state.richtextdescription3}
                                isEditMode={!this.state.Disabled}
                                onChange={(text) => {
                                  this.setState({ richtextdescription3: text });
                                  return text;
                                }}
                              /> */}
                              <textarea className="content3" name="example3"
                                onChange={(e) => this.setState({ richtextdescription3: e.target.value })} />
                            </td>
                            <td>
                              <select className="form-control" disabled={this.state.Disabled}
                                value={this.state.ImprestCheck3} onChange={(e) => this.setState({ ImprestCheck3: e.target.value })}>
                                <option value={''}>--Select--</option>
                                {this.state.CheckListDropdown.map((items) => {
                                  return (
                                    <option value={items.ID}>{items.CheckList}</option>
                                  )
                                })}
                              </select>
                            </td>
                          </tr>
                          <tr>
                            <td>4</td>
                            <td>
                              <p>Imprest payments are not batch processed in GSM.</p>
                              {/* <RichText value={this.state.richtextdescription4}
                                isEditMode={!this.state.Disabled}
                                onChange={(text) => {
                                  this.setState({ richtextdescription4: text });
                                  return text;
                                }}
                              /> */}
                              <textarea className="content4" name="example4"
                                onChange={(e) => this.setState({ richtextdescription4: e.target.value })} />
                            </td>
                            <td>
                              <select className="form-control" disabled={this.state.Disabled}
                                value={this.state.ImprestCheck4} onChange={(e) => this.setState({ ImprestCheck4: e.target.value })}>
                                <option value={''}>--Select--</option>
                                {this.state.CheckListDropdown.map((items) => {
                                  return (
                                    <option value={items.ID}>{items.CheckList}</option>
                                  )
                                })}
                              </select>
                            </td>
                          </tr>
                          <tr>
                            <td>5</td>
                            <td>
                              <p>Imprest ceiling is adequate.</p>
                              {/* <RichText value={this.state.richtextdescription5}
                                isEditMode={!this.state.Disabled}
                                onChange={(text) => {
                                  this.setState({ richtextdescription5: text });
                                  return text;
                                }}
                              /> */}
                              <textarea className="content5" name="example5"
                                onChange={(e) => this.setState({ richtextdescription5: e.target.value })} />
                            </td>
                            <td>
                              <select className="form-control" disabled={this.state.Disabled}
                                value={this.state.ImprestCheck5} onChange={(e) => this.setState({ ImprestCheck5: e.target.value })}>
                                <option value={''}>--Select--</option>
                                {this.state.CheckListDropdown.map((items) => {
                                  return (
                                    <option value={items.ID}>{items.CheckList}</option>
                                  )
                                })}
                              </select>
                            </td>
                          </tr>
                          <tr>
                            <td>6</td>
                            <td>
                              <p>Cash or bank balance was positive throughout the month and no overdraft was observed at any point during the month.</p>
                              {/* <RichText value={this.state.richtextdescription6}
                                isEditMode={!this.state.Disabled}
                                onChange={(text) => {
                                  this.setState({ richtextdescription6: text });
                                  return text;
                                }}
                              /> */}
                              <textarea className="content6" name="example6"
                                onChange={(e) => this.setState({ richtextdescription6: e.target.value })} />
                            </td>
                            <td>
                              <select className="form-control" disabled={this.state.Disabled}
                                value={this.state.ImprestCheck6} onChange={(e) => this.setState({ ImprestCheck6: e.target.value })}>
                                <option value={''}>--Select--</option>
                                {this.state.CheckListDropdown.map((items) => {
                                  return (
                                    <option value={items.ID}>{items.CheckList}</option>
                                  )
                                })}
                              </select>
                            </td>
                          </tr>
                          <tr>
                            <td>7</td>
                            <td>
                              <p>Cheques have been issued serially and Cheque Number(s)/P.I. Number(s) are mentioned appropriately for each of the transaction in GSM.</p>
                              {/* <RichText value={this.state.richtextdescription7}
                                isEditMode={!this.state.Disabled}
                                onChange={(text) => {
                                  this.setState({ richtextdescription7: text });
                                  return text;
                                }}
                              /> */}
                              <textarea className="content7" name="example7"
                                onChange={(e) => this.setState({ richtextdescription7: e.target.value })} />
                            </td>
                            <td>
                              <select className="form-control" disabled={this.state.Disabled}
                                value={this.state.ImprestCheck7} onChange={(e) => this.setState({ ImprestCheck7: e.target.value })}>
                                <option value={''}>--Select--</option>
                                {this.state.CheckListDropdown.map((items) => {
                                  return (
                                    <option value={items.ID}>{items.CheckList}</option>
                                  )
                                })}
                              </select>
                            </td>
                          </tr>
                          <tr>
                            <td>8</td>
                            <td>
                              <p>Cash-in-safe has been maintained within the insurance limit at all the times during the month.</p>
                              {/* <RichText value={this.state.richtextdescription8}
                                isEditMode={!this.state.Disabled}
                                onChange={(text) => {
                                  this.setState({ richtextdescription8: text });
                                  return text;
                                }}
                              /> */}
                              <textarea className="content8" name="example8"
                                onChange={(e) => this.setState({ richtextdescription8: e.target.value })} />
                            </td>
                            <td>
                              <select className="form-control" disabled={this.state.Disabled}
                                value={this.state.ImprestCheck8} onChange={(e) => this.setState({ ImprestCheck8: e.target.value })}>
                                <option value={''}>--Select--</option>
                                {this.state.CheckListDropdown.map((items) => {
                                  return (
                                    <option value={items.ID}>{items.CheckList}</option>
                                  )
                                })}
                              </select>
                            </td>
                          </tr>
                          <tr>
                            <td>9</td>
                            <td>
                              <p>Are proper handing over documents signed and attached to the Imprest Returns, if applicable?</p>
                              {/* <RichText value={this.state.richtextdescription9}
                                isEditMode={!this.state.Disabled}
                                onChange={(text) => {
                                  this.setState({ richtextdescription9: text });
                                  return text;
                                }}
                              /> */}
                              <textarea className="content9" name="example9"
                                onChange={(e) => this.setState({ richtextdescription9: e.target.value })} />
                            </td>
                            <td>
                              <select className="form-control" disabled={this.state.Disabled}
                                value={this.state.ImprestCheck9} onChange={(e) => this.setState({ ImprestCheck9: e.target.value })}>
                                <option value={''}>--Select--</option>
                                {this.state.CheckListDropdown.map((items) => {
                                  return (
                                    <option value={items.ID}>{items.CheckList}</option>
                                  )
                                })}
                              </select>
                            </td>
                          </tr>
                          <tr>
                            <td>10</td>
                            <td>
                              <p>Currency conversion greater than $100,000 is supported by competitive bidding process and cumulative currency conversion in excess of $500,000 per month is coordinated through BFO/SEARO.</p>
                              {/* <RichText value={this.state.richtextdescription10}
                                isEditMode={!this.state.Disabled}
                                onChange={(text) => {
                                  this.setState({ richtextdescription10: text });
                                  return text;
                                }}
                              /> */}
                              <textarea className="content10" name="example10"
                                onChange={(e) => this.setState({ richtextdescription10: e.target.value })} />
                            </td>
                            <td>
                              <select className="form-control" disabled={this.state.Disabled}
                                value={this.state.ImprestCheck10} onChange={(e) => this.setState({ ImprestCheck10: e.target.value })}>
                                <option value={''}>--Select--</option>
                                {this.state.CheckListDropdown.map((items) => {
                                  return (
                                    <option value={items.ID}>{items.CheckList}</option>
                                  )
                                })}
                              </select>
                            </td>
                          </tr>
                          <tr>
                            <td>11</td>
                            <td>
                              <p>There is no other comment to add.</p>
                              {/* <RichText value={this.state.richtextdescription11}
                                isEditMode={!this.state.Disabled}
                                onChange={(text) => {
                                  this.setState({ richtextdescription11: text });
                                  return text;
                                }}
                              /> */}
                              <textarea className="content11" name="example11"
                                onChange={(e) => this.setState({ richtextdescription11: e.target.value })} />
                            </td>
                            <td>
                              <select className="form-control" disabled={this.state.Disabled}
                                value={this.state.ImprestCheck11} onChange={(e) => this.setState({ ImprestCheck11: e.target.value })}>
                                <option value={''}>--Select--</option>
                                {this.state.CheckListDropdown.map((items) => {
                                  return (
                                    <option value={items.ID}>{items.CheckList}</option>
                                  )
                                })}
                              </select>
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                    <div className="clearfix"></div>
                    <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12 p-0">
                      <div className="form-group">
                        <label className="col-lg-4 col-md-4 col-sm-4 col-xs-12 pl0 pt15">Feedback</label>
                        <textarea className="form-control" rows={4}
                          value={this.state.Feedback}
                          disabled={this.state.Disabled}
                          onChange={(e) => this.setState({ Feedback: e.target.value })} />
                      </div>
                    </div>
                    <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12 p-0">
                      <div className="col-lg-4 col-md-4 col-sm-6 col-xs-12 p-0">
                        <div className="row">
                          <div className="form-group">
                            <label className="col-lg-4 col-md-4 col-sm-4 col-xs-12">Attachment(s):</label>
                            <div className="col-lg-8 col-md-8 col-sm-8 col-xs-12">
                              <input className="form-control" type="file" name="uploadfile" disabled={this.state.Disabled}
                                id="img" style={{ display: 'none' }} onChange={(e) => this.setButtonsEventHandlersForAttachment(e.target.files)} />
                              <label htmlFor="img" className="d-block"><i className="fa fa-upload upload-font"></i> Click to upload file</label>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                    <div className="col-sm-8 table-responsive p-0">
                      <table className="table table-bordered table-white">
                        <thead className="blue_bg">
                          <tr>
                            <th>S No.</th>
                            <th>File Name</th>
                            <th>Attached by</th>
                            <th>Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {this.state.FileArr.map((item, index) => {
                            const person = item.person;
                            return (
                              <tr>
                                <td>{index}</td>
                                <td><a href="#">{item.name}</a></td>
                                {this.state.Disabled && person ? (
                                  <>
                                    <td>{person}</td>
                                    <td style={{ fontSize: "18px" }}> <a href="#" style={{ pointerEvents: 'none' }}
                                      onClick={() => {
                                        let file: any[] = fileInfos.splice(fileInfos.findIndex(a => a.name === item.name), 1)
                                        this.setState({ FileArr: file })
                                        this.setState({ FileArr: fileInfos });
                                        console.log(fileInfos, this.state.FileArr);
                                        // fileInfos = this.state.FileArr;
                                      }}>
                                      <i className="fa fa-trash"></i></a> </td>
                                  </>
                                ) : (
                                  <>
                                    <td>{person ? person : this.props.context.pageContext.user.displayName}</td>
                                    <td style={{ fontSize: "18px" }}> <a href="#"
                                      onClick={() => {
                                        this.deleteFile(itemId, item.name)
                                        let file: any[] = fileInfos.splice(fileInfos.findIndex(a => a.name === item.name), 1)
                                        this.setState({ FileArr: file })
                                        this.setState({ FileArr: fileInfos });
                                        console.log(fileInfos, this.state.FileArr);
                                        // fileInfos = this.state.FileArr;
                                      }}>
                                      <i className="fa fa-trash"></i></a> </td>
                                  </>
                                )}

                              </tr>
                            )
                          })}
                        </tbody>
                      </table>
                    </div>
                    <div className="col-lg-12 col-md-12 col-sm-12 col-xs-12 mt20 text-center">
                      {
                        /// querystring id is passed and record is not completed mode (draft)
                        this.state.UpdateEditForm && !this.state.Disabled ? (
                          <>
                            <button type="button" className="btn btn-warning mr5"
                              onClick={() => this.updateTask(itemId, "Draft", "No")}>Save</button>
                            <button type="button" className="btn btn-success mr5"
                              onClick={() => this.updateTask(itemId, "Completed", "Yes")}>Submit</button>
                          </>
                        ) :
                          (
                            <>
                              <button type="button" className="btn btn-warning mr5" disabled={this.state.Disabled}
                                onClick={() => this.submitTask("Draft", "No")}>Save</button>
                              <button type="button" className="btn btn-success mr5" disabled={this.state.Disabled}
                                onClick={() => this.submitTask("Completed", "Yes")}>Submit</button>
                            </>
                          )
                      }
                      <button type="button" className="btn btn-danger mr5" onClick={this.printpdf}>Print</button>
                      <button type="button" className="btn btn-primary mr5" onClick={() => {
                        window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/DashboardList.aspx`;
                      }}>Close</button>
                      <button className="btn btn-default btn-lg" id="loader"><i className="fa fa-refresh fa-spin"></i> Loading..</button>

                    </div>
                  </div>
                </div>
              </div>
            </section>
          </div>
        </div>

      </>
    );

  }
}
