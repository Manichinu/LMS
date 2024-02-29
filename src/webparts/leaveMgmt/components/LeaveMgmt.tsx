import * as React from 'react';
import { useState, useEffect } from "react";
import styles from './LeaveMgmt.module.scss';
import { ILeaveMgmtProps } from './ILeaveMgmtProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from "@microsoft/sp-loader";
import * as $ from "jquery";
import "bootstrap-datepicker";
import ReactFileReader from 'react-file-reader';


import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "@pnp/sp/presets/all";

import { IAttachmentFileInfo, IItemAddResult, Web } from "@pnp/sp/presets/all";
import swal from "sweetalert";
import * as moment from "moment";

import { _Items } from '@pnp/sp/items/types';
import DatePicker, { DateObject } from "react-multi-date-picker";
import DatePanel from "react-multi-date-picker/plugins/date_panel";


var AttachmentURL = "";
var FileNameGenerated: string;
let inputFile: any = '';
var fileInfos: any[] = [];

//var fileInfos: IAttachmentFileInfo[] = [];
var CurrentUSERNAME = "";
let LRUploadedFiles = [];
//let AttachmentCopies = [];
let ItemId;
let StartdateArr = [];
let datesCollection: string[] = [];
var PreviousLeaveRequestDates: any[] = [];
var PreviousPermissionRequestDates = [];
var AllFileAttachmentURL = "";
var finalurl = "";
var Approver_Manager_Details: any = []
var format = 'DD/MM/YYYY';
var RestrictedHoliday: any;

export interface ILeaveMgmtState {
  items: any[];
  LeaveID: any;
  //EmployeeName:any;
  leavetype: any;
  Day: any;
  Time: any;
  Startdate: any;
  Enddate: any;
  Reason: any;
  //Attachments: any;
  //fileInfos: any[];
  uploadfiles: any[];
  CurrentUserName: string;
  CurrentUserDesignation: string;
  CurrentUserProfilePic: string;
  Gender: string;
  Email: string;
  AttachmentCopies: any[];
  leaveBindCopy: [],
  Appliedleaveitems: any[],
  IsAlreadyexist: boolean;
  dates: any[]
  EndDate: any;
  DatePickerDisable: boolean;
}

//let NewWeb = Web("https://tmxin.sharepoint.com/sites/POC/SPIP/");
//let NewWeb = Web("https://tmxin.sharepoint.com/sites/lms");
let NewWeb = Web("https://tmxin.sharepoint.com/sites/ER");
export default class LeaveMgmt extends React.Component<ILeaveMgmtProps, ILeaveMgmtState> {

  public constructor(props: ILeaveMgmtProps, stated: ILeaveMgmtState) {

    super(props);

    let leaveBindCopy = [];
    var fileArr = [];
    var CurrentUSERNAME = ""
    SPComponentLoader.loadCss(
      `https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css`
    );
    SPComponentLoader.loadCss(`https://fonts.googleapis.com`);
    SPComponentLoader.loadCss(`https://fonts.gstatic.com" crossorigin`);
    SPComponentLoader.loadCss(
      `https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet`
    );
    SPComponentLoader.loadCss(
      `https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css`
    );
    //newly added
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.7.1/css/bootstrap-datepicker3.min.css');
    //
    SPComponentLoader.loadScript(
      `https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js`
    );

    SPComponentLoader.loadScript(
      `https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.js`
    );


    SPComponentLoader.loadCss(
      `https://tmxin.sharepoint.com/sites/ER/SiteAssets/LeavePortal/css/style.css?v=1.14`
    );



    this.state = {
      items: [],
      leavetype: "",
      Day: 0,
      Time: null,
      Startdate: null,
      Enddate: null,
      Reason: "",
      // Attachments: null,
      LeaveID: null,
      //fileInfos: [],
      AttachmentCopies: [],
      uploadfiles: null,
      CurrentUserName: "",
      CurrentUserDesignation: "",
      CurrentUserProfilePic: "",
      Gender: "",
      Email: "",
      leaveBindCopy: [],
      Appliedleaveitems: [],
      IsAlreadyexist: false,
      dates: [],
      EndDate: "",
      DatePickerDisable: false

    };
    //this.handleChange = this.handleChange.bind(this);
  }



  public LoadDatePicker() {
    ($("#txt-Startdate") as any).datepicker({
      startDate: new Date(),
      format: 'dd-mm-yyyy',
      daysOfWeekDisabled: [0, 6],

      clearBtn: true,
      autoclose: true,
      multidate: true, // Enable multidate selection
      multidateSeparator: ', ' // Separator for displaying selected dates
    });
    ($("#txt-Enddate") as any).datepicker({
      startDate: new Date(),
      format: 'dd-mm-yyyy',
      daysOfWeekDisabled: [0, 6],
      todayHighlight: true,
      clearBtn: true,
      autoclose: true
    });
  }

  public logout() {

    localStorage.clear();
    window.location.href = `https://login.windows.net/common/oauth2/logout`;

  }


  public GetCausalleavebalance(email: any) {

    var reactHandler = this;
    let currentYear = new Date().getFullYear()
    let nextYear = currentYear + 1
    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");

    NewWeb.lists.getByTitle("BalanceCollection").items.select("Id", "*", "CasualLeaveBalance", "EmployeeEmail").filter(`EmployeeEmail eq '${email}' and StartDate eq '01/04/${currentYear}' and EndDate eq '31/03/${nextYear}'`).get()

      .then((items) => {

        if (items.length != 0) {
          for (var i = 0; i < items.length; i++) {
            if (items[i].CasualLeaveBalance == 0) {
              swal({
                text: "You have availed the limit of the choosen leave",
                icon: "error",
              });
            }
            else {

              this.Addtolist();
            }

          }

        }
      });


  }
  public CheckValidationforExistingleave(email: any) {
    datesCollection = [];
    var reactHandler = this;
    // var startdate = $("#txt-Startdate").val();
    //var enddate = $("#txt-Enddate").val();
    //var filterquery = `${this.state.Email}' and '${startdate}' ge startdate and '${enddate}' le enddate`
    var url = `https://tmxin.sharepoint.com/sites/ER/_api/web/lists/getbytitle('LeaveRequest')/items?$select=StartDate,EndDate,EmployeeEmail&$filter('Author/EmployeeEmail eq '${email}'')`;

    $.ajax({
      url: url,
      type: "GET",
      async: false,
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        // console.log(resultData);
        debugger;
        {/* for (var i = 0; i < resultData.d.results.length; i++) {

          var sdate = resultData.d.results[i].startdate;
          var edate = resultData.d.results[i].enddate;
          var mailid = resultData.d.results[i].EmployeeEmail;

        }*/}

        reactHandler.setState({
          Appliedleaveitems: resultData.d.results
        });


        for (var i = 1; i < resultData.d.results.length; i++) {
          var sdate = resultData.d.results[i].startdate;
          var edate = resultData.d.results[i].enddate;
          let date = [];

          for (var m = moment(sdate); m.isSameOrBefore(edate); m.add(1, 'days')) {
            datesCollection.push(m.format('YYYY-MM-DD'));
          }
          console.log(datesCollection);

          // datesCollection.push(`${inbetweendates.getDate() + i}/${inbetweendates.getMonth() + 1}/${inbetweendates.getFullYear()}`)

          //  console.log(datesCollection);


          {/* if (resultData.d.results[i].startdate == "startdate") {
            StartdateArr.push(resultData.d.results[i]);
          }*/}

        }
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });
  }
  public GetPreviousLeaveRequestDates(email: any) {

    //var filterquery = `EmployeeEmail eq '${this.state.Email}' and '${permissionDate}' ge startdate and '${permissionDate}' le enddate`
    //// and enddate ge '${moment().format("DD-MM-YYYY")}'
    var filterquery = `EmployeeEmail eq '${email}' and Status ne 'Rejected'`
    NewWeb.lists.getByTitle("LeaveRequest").items.select("StartDate", "EndDate", "EmployeeEmail", "Status").filter(filterquery).orderBy("Created", false).get().then((response: any): void => {
      if (response.length != 0) {
        let i;
        for (i = 0; i < response.length;) {
          var From = response[i].StartDate;
          console.log(From);

          var To = response[i].EndDate;
          console.log(To);

          var tempFromDate = moment(From).format("YYYY-MM-DD");
          console.log(tempFromDate);
          var tempToDate = moment(To).format("YYYY-MM-DD");
          console.log(tempToDate);
          var dateList = this.getDaysBetweenDates(moment(tempFromDate), moment(tempToDate), response[i].Status);
          console.log("dateList LeaveRequest: " + dateList);
          i++;
        }
      }
    });
  }
  public GetPreviousPermissionRequestDates(email: any) {

    // var filterquery = `EmployeeEmail eq '${this.state.Email}' and '${permissionDate}' ge startdate and '${permissionDate}' le enddate`
    debugger;
    var filterquery = `EmployeeEmail eq '${email}' and Status ne 'Rejected'`// and timefromwhen ge '${moment().format("DD-MM-YYYY")}'`
    debugger;
    //  var filterquery = `Author/EmployeeEmail eq '${email}'and timefromwhen ge '${moment().format("DD-MM-YYYY")}'`
    NewWeb.lists.getByTitle("EmployeePermission").items.select("timefromwhen", "EmployeeEmail", "Status").filter(filterquery).orderBy("Created", false).get().then((response: any): void => {
      if (response.length != 0) {
        let i;
        for (i = 0; i < response.length;) {
          var From = response[i].timefromwhen;

          var tempFromDate = moment(From, "DD-MM-YYYYTHH:mm").format('DD-MM-YYYY');

          var tempToDate = moment(From, "DD-MM-YYYYTHH:mm").format("DD-MM-YYYY");


          var dateList = this.getDaysBetweenDates(moment(tempFromDate), moment(tempToDate), response[i].Status);
          console.log("dateList PermissionRequest: " + dateList);
          i++;
        }
      }
    });
  }
  public Checkalreadyinleave(email: any) {
    var startdate = $("#txt-Startdate").val();
    var enddate = $("#txt-Enddate").val();
    let Status = true;


    var filterquery = `EmployeeEmail eq '${email}' and '${startdate}' ge StartDate and '${enddate}' le EndDate`
    debugger;


    NewWeb.lists.getByTitle("LeaveRequest").items.select("Id", "StartDate", "EndDate").filter(filterquery).orderBy("Created", false).get().then((response: any): void => {
      if (response.length != 0) {
        console.log(response);
        this.setState({
          IsAlreadyexist: true
        });
        Status = false;
      } else {
        var filterquery1 = `EmployeeEmail eq '${this.state.Email}' and  timefromwhen eq '${startdate}' and '${enddate}'`
        NewWeb.lists.getByTitle("EmployeePermission").items.select("Id", "timefromwhen", "PermissionOn").filter(filterquery1).orderBy("Created", false).get().then((response: any): void => {
          if (response.length != 0) {
            this.setState({
              IsAlreadyexist: true
            });
            Status = false;
          }
          else {
            this.setState({

              IsAlreadyexist: false

            });
            Status = true;

          }
        });

      }

    })

    return Status;

  }
  public GetSickleavebalance(email: any) {

    var reactHandler = this;
    let currentYear = new Date().getFullYear()
    let nextYear = currentYear + 1
    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");

    NewWeb.lists.getByTitle("BalanceCollection").items.select("Id", "*", "SickLeaveBalance", "EmployeeEmail").filter(`EmployeeEmail eq '${email}' and StartDate eq '01/04/${currentYear}' and EndDate eq '31/03/${nextYear}'`).get()

      .then((items) => {

        if (items.length != 0) {
          for (var i = 0; i < items.length; i++) {
            if (items[i].SickLeaveBalance == 0) {
              swal({
                text: "You have availed the limit of the choosen leave",
                icon: "error",
              });
            }
            else {
              this.Addtolist();
            }

          }

        }
      });


  }
  public GetEarnedleavebalance(email: any) {

    var reactHandler = this;
    let currentYear = new Date().getFullYear()
    let nextYear = currentYear + 1
    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");

    NewWeb.lists.getByTitle("BalanceCollection").items.select("Id", "*", "EarnedLeaveBalance", "EmployeeEmail").filter(`EmployeeEmail eq '${email}' and StartDate eq '01/04/${currentYear}' and EndDate eq '31/03/${nextYear}'`).get()

      .then((items) => {

        if (items.length != 0) {
          for (var i = 0; i < items.length; i++) {
            if (items[i].EarnedLeaveBalance == 0) {
              swal({
                text: "You have availed the limit of the choosen leave",
                icon: "error",
              });
            }
            else {
              this.Addtolist();
            }
          }

        }
      });


  }
  public GetMaternityleavebalance(email: any) {

    var reactHandler = this;
    let currentYear = new Date().getFullYear()
    let nextYear = currentYear + 1
    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");

    NewWeb.lists.getByTitle("BalanceCollection").items.select("Id", "*", "MaternityLeaveBalance", "EmployeeEmail").filter(`EmployeeEmail eq '${email}' and StartDate eq '01/04/${currentYear}' and EndDate eq '31/03/${nextYear}'`).get()

      .then((items) => {

        if (items.length != 0) {
          for (var i = 0; i < items.length; i++) {
            if (items[i].MaternityLeaveBalance == 0) {
              swal({
                text: "You have availed the limit of the choosen leave",
                icon: "error",
              });
            }
            else {
              this.Addtolist();
            }
          }

        }
      });


  }
  public GetOtherleavebalance(email: string) {

    var reactHandler = this;
    let currentYear = new Date().getFullYear()
    let nextYear = currentYear + 1
    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");

    NewWeb.lists.getByTitle("BalanceCollection").items.select("Id", "*", "OtherLeaveBalance", "EmployeeEmail").filter(`EmployeeEmail eq '${email}' and StartDate eq '01/04/${currentYear}' and EndDate eq '31/03/${nextYear}'`).get()

      .then((items) => {

        if (items.length != 0) {
          for (var i = 0; i < items.length; i++) {
            if (items[i].OtherLeaveBalance == 0) {
              swal({
                text: "You have availed the limit of the choosen leave",
                icon: "error",
              });
            }
            else {
              this.Addtolist();
            }

          }

        }
      });


  }
  public GetPaternityleavebalance(email: any) {

    var reactHandler = this;
    let currentYear = new Date().getFullYear()
    let nextYear = currentYear + 1
    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");

    NewWeb.lists.getByTitle("BalanceCollection").items.select("Id", "*", "PaternityLeaveBalance", "EmployeeEmail").filter(`EmployeeEmail eq '${email}' and StartDate eq '01/04/${currentYear}' and EndDate eq '31/03/${nextYear}'`).get()

      .then((items) => {

        if (items.length != 0) {
          for (var i = 0; i < items.length; i++) {
            if (items[i].PaternityLeaveBalance == 0) {
              swal({
                text: "You have availed the limit of the choosen leave",
                icon: "error",
              });
            }
            else {
              this.Addtolist();
            }

          }

        }
      });

  }
  public GetCausalleaveExhaustbalance(email: string, Days: number) {

    var reactHandler = this;
    let currentYear = new Date().getFullYear()
    let nextYear = currentYear + 1
    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");

    NewWeb.lists.getByTitle("BalanceCollection").items.select("Id", "*", "CasualLeaveBalance", "CasualLeave", "EmployeeEmail").filter(`EmployeeEmail eq '${email}' and StartDate eq '01/04/${currentYear}' and EndDate eq '31/03/${nextYear}'`).get()

      .then((items) => {

        if (items.length != 0) {
          for (var i = 0; i < items.length; i++) {
            if (items[i].CasualLeaveBalance <= Days) {
              swal({
                text: "Maximum limit of casual leave is 12",
                icon: "error",
              });
            }
            else {
              this.Addtolist();
            }

          }

        }
      });
  }
  public GetPaternityleaveExhaustbalance(email: string, Days: number) {

    var reactHandler = this;
    let currentYear = new Date().getFullYear()
    let nextYear = currentYear + 1
    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");

    NewWeb.lists.getByTitle("BalanceCollection").items.select("Id", "*", "PaternityLeaveBalance", "PaternityLeave", "EmployeeEmail").filter(`EmployeeEmail eq '${email}' and StartDate eq '01/04/${currentYear}' and EndDate eq '31/03/${nextYear}'`).get()

      .then((items) => {

        if (items.length != 0) {
          for (var i = 0; i < items.length; i++) {
            if (items[i].PaternityLeaveBalance <= Days) {
              swal({
                text: "Maximum limit of paternity leave is 40",
                icon: "error",
              });
            }
            else {
              this.Addtolist();
            }

          }

        }
      });
  }
  public GetMaternityleaveExhaustbalance(email: string, Days: number) {

    var reactHandler = this;
    let currentYear = new Date().getFullYear()
    let nextYear = currentYear + 1
    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");

    NewWeb.lists.getByTitle("BalanceCollection").items.select("Id", "*", "MaternityLeaveBalance", "MaternityLeave", "EmployeeEmail").filter(`EmployeeEmail eq '${email}' and StartDate eq '01/04/${currentYear}' and EndDate eq '31/03/${nextYear}'`).get()

      .then((items) => {

        if (items.length != 0) {
          for (var i = 0; i < items.length; i++) {
            if (items[i].MaternityLeaveBalance <= Days) {
              swal({
                text: "Maximum limit of maternity leave is 130",
                icon: "error",
              });
            }
            else {
              this.Addtolist();
            }

          }

        }
      });
  }

  public clearerror() {
    $("#divErrorText").empty();
    $("#divErrorText").hide();
    var Leavetype = $("#ddl-leavetype").val();
    if (Leavetype == "Restricted Leave") {
      $("#comp_off_date").show()
      RestrictedHoliday = true
    } else {
      $("#comp_off_date").hide()
      RestrictedHoliday = false
      this.setState({
        dates: []
      })
    }
  }
  public GetEarnedleaveExhaustbalance(email: string, Days: number) {

    var reactHandler = this;
    let currentYear = new Date().getFullYear()
    let nextYear = currentYear + 1
    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");

    NewWeb.lists.getByTitle("BalanceCollection").items.select("Id", "*", "EarnedLeaveBalance", "EarnedLeave", "EmployeeEmail").filter(`EmployeeEmail eq '${email}' and StartDate eq '01/04/${currentYear}' and EndDate eq '31/03/${nextYear}'`).get()

      .then((items) => {

        if (items.length != 0) {
          for (var i = 0; i < items.length; i++) {
            if (items[i].EarnedLeaveBalance <= Days) {
              swal({
                text: "Maximum limit of earned leave is 12",
                icon: "error",
              });
            }
            else {
              this.Addtolist();
            }

          }

        }
      });
  }
  public GetSickleaveExhaustbalance(email: string, Days: number) {

    var reactHandler = this;
    let currentYear = new Date().getFullYear()
    let nextYear = currentYear + 1
    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");

    NewWeb.lists.getByTitle("BalanceCollection").items.select("Id", "*", "SickLeaveBalance", "SickLeave", "EmployeeEmail").filter(`EmployeeEmail eq '${email}' and StartDate eq '01/04/${currentYear}' and EndDate eq '31/03/${nextYear}'`).get()

      .then((items) => {

        if (items.length != 0) {
          for (var i = 0; i < items.length; i++) {
            if (items[i].SickLeaveBalance <= Days) {
              swal({
                text: "Maximum limit of sick leave is 12",
                icon: "error",

              })

            }
            else {
              this.Addtolist();
            }

          }

        }
      });
  }

  public disableweekends() {
    // disable weekends


    const dWeekends = (current: { day: () => number; }) => {
      return current.day() !== 0 && current.day() !== 6;
    }
  }
  public disablepastdates() {
    var date = new Date().toISOString().slice(0, 10);

    //To restrict past date

    $('#txt-Startdate').attr('min', date);
    $('#txt-Enddate').attr('min', date);
  }

  public componentDidMount() {
    // this.LoadDatePicker();
    $(".opttime").prop('disabled', true);
    $("#ddl-full-Day").prop("checked", true);
    this.GenerateLeaveID();

    this.GetCurrentUserDetails();


    var now = new Date();
    var day = ("0" + now.getDate()).slice(-2);
    var month = ("0" + (now.getMonth() + 1)).slice(-2);
    var today = now.getFullYear() + "-" + (month) + "-" + (day);


    $('#txt-Startdate').val(today);
    $('#txt-Enddate').val(today);
    $('#txt-Startdate').attr('min', today);
    $('#txt-Enddate').attr('min', today);
    var CompensateMinDate = moment($("#txt-Enddate").val(), "YYYY-MM-DD").add(1, 'days').format("YYYY-MM-DD")
    this.setState({ EndDate: CompensateMinDate })

    const url: any = new URL(window.location.href);
    const LeaveType = url.searchParams.get("type");
    setTimeout(() => {
      if (LeaveType == "TMX001") {

        $("#ddl-leavetype").val("Earned Leave");

        $("#ddl-leavetype").prop('disabled', true);

        var SelValLegth = $("#ddl-leavetype").val();
        if (SelValLegth != "") {
          $("#ddl-leave-type-label").addClass("disabled");
        }
      }
      else if (LeaveType == "TMX002") {

        $("#ddl-leavetype").val("Casual Leave");

        $("#ddl-leavetype").prop('disabled', true);

        var SelValLegth = $("#ddl-leavetype").val();
        if (SelValLegth != "") {
          $("#ddl-leave-type-label").addClass("disabled");
        }
      }
      else if (LeaveType == "TMX003") {

        $("#ddl-leavetype").val("Sick Leave");

        $("#ddl-leavetype").prop('disabled', true);
        var SelValLegth = $("#ddl-leavetype").val();
        if (SelValLegth != "") {
          $("#ddl-leave-type-label").addClass("disabled");
        }
      }
      else if (LeaveType == "TMX004") {

        $("#ddl-leavetype").val("Maternity Leave");

        $("#ddl-leavetype").prop('disabled', true);
        var SelValLegth = $("#ddl-leavetype").val();
        if (SelValLegth != "") {
          $("#ddl-leave-type-label").addClass("disabled");
        }
      }
      else if (LeaveType == "TMX005") {
        ;
        $("#ddl-leavetype").val("Paternity Leave");

        $("#ddl-leavetype").prop('disabled', true);
        var SelValLegth = $("#ddl-leavetype").val();
        if (SelValLegth != "") {
          $("#ddl-leave-type-label").addClass("disabled");
        }
      }
      else if (LeaveType == "TMX006") {

        $("#ddl-leavetype").val("Unpaid Leave");

        $("#ddl-leavetype").prop('disabled', true);
        var SelValLegth = $("#ddl-leavetype").val();
        if (SelValLegth != "") {
          $("#ddl-leave-type-label").addClass("disabled");
        }
      }
    }, 800);

  }



  public GetCurrentUserDetails() {

    var reactHandler = this;

    $.ajax({

      url: `${reactHandler.props.siteurl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`,

      type: "GET",

      headers: { 'Accept': 'application/json; odata=verbose;' },

      success: function (resultData) {

        var email = resultData.d.Email;

        var Name = resultData.d.DisplayName;

        var Designation = resultData.d.Title;
        var gender = resultData.d.Streetaddress;

        reactHandler.setState({

          CurrentUserName: Name,

          CurrentUserDesignation: Designation,
          Gender: gender,
          Email: email,

          CurrentUserProfilePic: `${reactHandler.props.siteurl}/_layouts/15/userphoto.aspx?size=l&username=${email}`

        });
        reactHandler.Get_CorrespondingApprover(resultData.d.Email)
        reactHandler.GetPreviousLeaveRequestDates(email);
        reactHandler.GetPreviousPermissionRequestDates(email);

        // reacthandler.CheckValidationforExistingleave(email); //Get User data from list and bind it in form 
      },

      error: function (jqXHR, textStatus, errorThrown) {

      }

    });

  }

  public Get_CorrespondingApprover(EmployeeEmailid: any) {
    var currentYear = new Date().getFullYear()
    let nextYear = currentYear + 1
    NewWeb.lists.getByTitle("BalanceCollection").items.select("ID", "*", "CasualLeaveBalance", "EmployeeEmail", "Manager/Title", "Manager/EMail").expand("Manager").filter(`EmployeeEmail eq '${EmployeeEmailid}' and StartDate eq '01/04/${currentYear}' and EndDate eq '31/03/${nextYear}'`).get()
      .then((result) => {
        if (result.length != 0) {
          console.log(result);

          Approver_Manager_Details.push({
            ApproverName: result[0].Manager.Title,
            ApproverEmail: result[0].Manager.EMail
          })

          console.log(Approver_Manager_Details)
        }
      })
  }


  public leavetypevalidation() {
    if (this.LeaveformValidation()) {
      $("#divErrorText").empty();
      var leavetype = $("#ddl-leavetype").val();
      var Startdate = $("#txt-Startdate").val();
      var EndDate = $("#txt-Enddate").val();

      var start = moment(Startdate, "YYYY-MM-DD");
      var end = moment(EndDate, "YYYY-MM-DD");
      var Days = moment.duration(end.diff(start)).asDays();
      if (leavetype == "Casual Leave") {

        // this.GetCausalleavebalance(this.state.Email);
        this.GetCausalleaveExhaustbalance(this.state.Email, Days);
        return false;
      } else {
        if (leavetype == "Sick Leave") {

          // this.GetSickleavebalance(this.state.Email);
          this.GetSickleaveExhaustbalance(this.state.Email, Days);
          return false;
        } else {
          if (leavetype == "Paternity Leave") {

            //   this.GetPaternityleavebalance(this.state.Email);
            this.GetPaternityleaveExhaustbalance(this.state.Email, Days);

            return false;
          } else {
            if (leavetype == "Maternity Leave") {

              // this.GetMaternityleavebalance(this.state.Email);
              this.GetMaternityleaveExhaustbalance(this.state.Email, Days);

              return false;
            } else {
              if (leavetype == "Earned Leave") {

                //this.GetEarnedleavebalance(this.state.Email);
                this.GetEarnedleaveExhaustbalance(this.state.Email, Days);
                return false;
              } else {
                if (leavetype == "Unpaid Leave") {

                  this.GetOtherleavebalance(this.state.Email);

                  return false;
                }

              }
            }
          }
        }
      }
      if (leavetype == "Restricted Leave") {
        this.Addtolist();
      }

    }
  }
  public Addtolist() {
    $("#divErrorText").empty();
    $("#divErrorText").hide();
    if (this.LeaveformValidation()) {

      var now = new Date();
      var day = ("0" + now.getDate()).slice(-2);
      var month = ("0" + (now.getMonth() + 1)).slice(-2);
      var today = now.getFullYear() + "-" + (month) + "-" + (day);

      var leavetype = $("#ddl-leavetype").val();
      var Day = $('input[name="optradio"]:checked').val();
      var Time = $('input[name="optradio1"]:checked').val();
      var startdate = $("#txt-Startdate").val();
      var enddate = $("#txt-Enddate").val();
      var Reason = $("#txt-reason").val();
      var fupload: any = document.querySelector(".leave-file-upload");

      var start = moment(startdate, "YYYY-MM-DD");
      var end = moment(enddate, "YYYY-MM-DD");
      var CompOfff: any = []
      this.state.dates.map((item) => {
        CompOfff.push(moment(item, "DD/MM/YYYY").format("DD-MM-YYYY"))
      })
      var CompOffDates = CompOfff.join(",")

      //Difference in number of days
      var Days = moment.duration(end.diff(start)).asDays();
      console.log(Days);

      if (Day == "Half Day") {
        var bool1 = this.checkIsafter(end, start);  // true

        if (bool1 == true) {

          swal({

            title: "End Date should be same as Start Date",

            icon: "error"

          });
          return false;

        }
        else {


          var bool = this.checkIsSameOrBefore(end, start);  // true

          if (bool == true) {

            swal({

              title: "End Date should be greater than Start Date",

              icon: "error"

            });

            return false;

          } else {
            var leavestart = moment(startdate, "YYYY-MM-DD").format('YYYY-MM-DD');
            var leaveend = moment(enddate, "YYYY-MM-DD").format('YYYY-MM-DD');

            if (this.isInArray(PreviousLeaveRequestDates, leavestart) == false) {//6 not found
              if (this.isInArray(PreviousLeaveRequestDates, leaveend) == false) {//6 not found
                if (this.LeaveformValidation()) {

                  var Days = 0.5;
                  NewWeb.lists.getByTitle("LeaveRequest").items.add({

                    LeaveType: leavetype,
                    Day: Day,
                    Time: Time,
                    StartDate: startdate,
                    EndDate: enddate,
                    Reason: Reason,
                    Requester: this.state.CurrentUserName,
                    AppliedDate: today,
                    Days: Days,
                    EmployeeEmail: this.state.Email,
                    RequestSessionMasterID: this.state.LeaveID,
                    Approver: Approver_Manager_Details[0].ApproverName,
                    ApproverEmail: Approver_Manager_Details[0].ApproverEmail,
                    CompOff: CompOffDates
                    // LRAttachments: finalurl
                  })
                    .then((item) => {

                      let ID = item.data.Id;
                      NewWeb.lists.getByTitle("LeaveRequest").items.getById(ID).attachmentFiles.addMultiple(this.state.AttachmentCopies).then(() => {
                        /*  NewWeb.lists.getByTitle("LeaveRequest").items.getById(ID).attachmentFiles.addMultiple(this.state.AttachmentCopies).then(() => {*/
                        swal({
                          text: "Leave applied successfully!",
                          icon: "success",
                        }).then(() => {
                          setTimeout(() => {
                            location.href = "https://tmxin.sharepoint.com/sites/ER/SitePages/Dashboard.aspx?env=WebView";
                          }, 3000);

                        });
                      });

                    });

                }
              }
              else {

                swal({

                  text: "Already leave request taken on selected date",
                  icon: "error"
                });
              }
            }
            else {

              swal({

                text: "Already leave request taken on selected date",
                icon: "error"
              });
            }
          }
        }
      }
      else {
        if (Day == "Full Day") {

          var now = new Date();
          var day = ("0" + now.getDate()).slice(-2);
          var month = ("0" + (now.getMonth() + 1)).slice(-2);
          var today = now.getFullYear() + "-" + (month) + "-" + (day);

          var leavetype = $("#ddl-leavetype").val();
          var Day = $('input[name="optradio"]:checked').val();
          var Time = $('input[name="optradio1"]:checked').val();
          var startdate = $("#txt-Startdate").val();
          var enddate = $("#txt-Enddate").val();
          var Reason = $("#txt-reason").val();
          var fupload: any = document.querySelector(".leave-file-upload");

          var start = moment(startdate, "YYYY-MM-DD");
          var end = moment(enddate, "YYYY-MM-DD");

          var bool1 = this.checkIsSame(end, start);  // true

          if (bool1 == true) {
            var bool = this.checkIsSameOrBefore(end, start);  // true
            var Days = 1;
            if (bool == true) {

              swal({

                title: "End Date should be greater than Start Date",

                icon: "error"

              });

              return false;

            } else {

              var leavestart = moment(startdate, "YYYY-MM-DD").format('YYYY-MM-DD');
              var leaveend = moment(enddate, "YYYY-MM-DD").format('YYYY-MM-DD');
              if (this.isInArray(PreviousLeaveRequestDates, leavestart) == false) {//6 not found
                if (this.isInArray(PreviousLeaveRequestDates, leaveend) == false) {//6 not found
                  if (this.LeaveformValidation()) {

                    Days = 1;


                    // var Days = moment.duration(end.diff(start)).asDays();
                    NewWeb.lists.getByTitle("LeaveRequest").items.add({

                      LeaveType: leavetype,
                      Day: Day,
                      Time: Time,
                      StartDate: startdate,
                      EndDate: enddate,
                      Reason: Reason,
                      Requester: this.state.CurrentUserName,
                      AppliedDate: today,
                      Days: Days,
                      EmployeeEmail: this.state.Email,
                      RequestSessionMasterID: this.state.LeaveID,
                      Approver: Approver_Manager_Details[0].ApproverName,
                      ApproverEmail: Approver_Manager_Details[0].ApproverEmail,
                      CompOff: CompOffDates

                    })
                      .then((item) => {

                        let ID = item.data.Id;
                        NewWeb.lists.getByTitle("LeaveRequest").items.getById(ID).attachmentFiles.addMultiple(this.state.AttachmentCopies).then(() => {
                          /*  NewWeb.lists.getByTitle("LeaveRequest").items.getById(ID).attachmentFiles.addMultiple(this.state.AttachmentCopies).then(() => {*/
                          swal({
                            text: "Leave applied successfully!",
                            icon: "success",
                          }).then(() => {
                            setTimeout(() => {
                              location.href = "https://tmxin.sharepoint.com/sites/ER/SitePages/Dashboard.aspx?env=WebView";
                            }, 3000);

                          });
                        });

                      });

                  }
                }
                else {

                  swal({

                    text: "Already leave request taken on selected date",
                    icon: "error"
                  });
                }
              }
              else {

                swal({

                  text: "Already leave request taken on selected date",
                  icon: "error"
                });
              }
            }
            return false;
          }

          else {

            var bool = this.checkIsSameOrBefore(end, start);  // true

            if (bool == true) {

              swal({

                title: "End Date should be greater than Start Date",

                icon: "error"

              });

              return false;

            } else {
              var leavestart = moment(startdate, "YYYY-MM-DD").format('YYYY-MM-DD');
              var leaveend = moment(enddate, "YYYY-MM-DD").format('YYYY-MM-DD');
              if (this.isInArray(PreviousLeaveRequestDates, leavestart) == false) {//6 not found
                if (this.isInArray(PreviousLeaveRequestDates, leaveend) == false) {//6 not found
                  if (this.LeaveformValidation()) {

                    //var Days = moment.duration(end.diff(start)).asDays();
                    var Days = moment.duration(end.diff(start)).add(1, 'days').asDays();

                    //var Days = 1;

                    // var Days = moment.duration(end.diff(start)).asDays();
                    var Days = moment.duration(end.diff(start)).add(1, 'days').asDays();
                    NewWeb.lists.getByTitle("LeaveRequest").items.add({

                      LeaveType: leavetype,
                      Day: Day,
                      Time: Time,
                      StartDate: startdate,
                      EndDate: enddate,
                      Reason: Reason,
                      Requester: this.state.CurrentUserName,
                      AppliedDate: today,
                      Days: Days,
                      EmployeeEmail: this.state.Email,
                      RequestSessionMasterID: this.state.LeaveID,
                      Approver: Approver_Manager_Details[0].ApproverName,
                      ApproverEmail: Approver_Manager_Details[0].ApproverEmail,
                      CompOff: CompOffDates


                    })
                      .then((item) => {

                        let ID = item.data.Id;
                        NewWeb.lists.getByTitle("LeaveRequest").items.getById(ID).attachmentFiles.addMultiple(this.state.AttachmentCopies).then(() => {
                          // NewWeb.lists.getByTitle("LeaveRequest").fields.addUrl("LRAttachments", {DisplayFormat: UrlFieldFormatType.Hyperlink});
                          /*  NewWeb.lists.getByTitle("LeaveRequest").items.getById(ID).attachmentFiles.addMultiple(this.state.AttachmentCopies).then(() => {*/
                          swal({
                            text: "Leave applied successfully!",
                            icon: "success",
                          }).then(() => {
                            setTimeout(() => {
                              location.href = "https://tmxin.sharepoint.com/sites/ER/SitePages/Dashboard.aspx?env=WebView";
                            }, 3000);

                          });
                        });

                      });

                  }
                }
                else {

                  swal({

                    text: "Already leave request taken on selected date",
                    icon: "error"
                  });
                }
              }
              else {

                swal({

                  text: "Already leave request taken on selected date",
                  icon: "error"
                });
              }

            }
          }

        }



      }


    }
  }

  public isInArray(PreviousLeaveRequestDates: any, value: string) {
    var DateStatus = false;
    // return (PreviousLeaveRequestDates.find((item: any) => { return item.Date == value && item.Status == "Cancelled" }) || []).length > 0;
    PreviousLeaveRequestDates.map((item: any) => {
      if (item.Date == value && (item.Status == "Approved" || item.Status == "Pending")) {
        DateStatus = true;
        return;
      }
    })
    return DateStatus
  }

  public getDaysBetweenDates(startDate: moment.Moment, endDate: moment.Moment, Status: any) {
    var now = startDate.clone();
    while (now.isSameOrBefore(endDate)) {

      PreviousLeaveRequestDates.push({ Date: now.format('YYYY-MM-DD'), Status: Status });

      now.add(1, 'days').format('DD/MM/YYYY');
    }
    return PreviousLeaveRequestDates;
  };
  public TriggerAttachment() {
    var reactHandler = this;
    reactHandler.setState({
      AttachmentCopies: []

    });
    NewWeb.lists.getByTitle("LeaveRequest").items.select("*").filter(`ID eq '${this.state.LeaveID}'`).expand("File").get().then((items) => {

      if (items.length != 0) {

        reactHandler.setState({
          AttachmentCopies: items
        });
      }
    });
  }

  public uploadfilestodocuments() {
    var fileArr = [];
    var CurrentTime;

    let myfile = (document.querySelector(".leave-file-upload") as HTMLInputElement).files.length;


    for (var j = 0; j < myfile; j++) {
      let fileVal = (document.querySelector(".leave-file-upload") as HTMLInputElement).files[j];
      fileArr.push(fileVal);
    }

    for (var i = 0; i < fileArr.length; i++) {
      CurrentTime = moment().format("DMYYYYHMS"); //1110202191045
      var tempfilename = fileArr[i].name.split(".");
      FileNameGenerated = tempfilename[0] + "-" + CurrentTime + "." + tempfilename[1] + "";
      //  FileNameGenerated = tempfilename[0] + "." + tempfilename[1] + "";
      NewWeb.getFolderByServerRelativeUrl(this.props.context.pageContext.web.serverRelativeUrl + "/LeaverequestUploads").files.add(FileNameGenerated, fileArr[i], true).then((data) => {
        data.file.getItem().then((item) => {
          //  AllFileAttachmentURL += "" + "https://tmxin.sharepoint.com/" +data.data.ServerRelativeUrl;// + "|";
          AllFileAttachmentURL += "" + data.data.ServerRelativeUrl;// + "|";
          //get value



          //  this.AddtoListItem(AllFileAttachmentURL, FileNameGenerated);
          LRUploadedFiles.push(data);

          item.update({
            RequestSessionMasterID: this.state.LeaveID,

          }).then((myupdate) => {
            //  this.TriggerAttachment();
            console.log("File uploaded sucessfully : " + i + "");

          });

        });
      }).catch((error) => {
        console.log(error);
      });
    }
  }
  public async GenerateLeaveID() {
    var LeaveID;
    const list = NewWeb.lists
      .getByTitle("LeaveRequest")
      .items.select("ID", "RequestSessionMasterID")
      .top(1)
      .orderBy("Created", false)
      .get()
      .then((items) => {
        if (items.length == 0) {
          LeaveID = "0" + moment().format("DDMMYYYYHMS") + "0000";

        } else {
          let num = parseInt(items[0].ID) + 1;
          LeaveID = "0" + moment().format("DDMMYYYYHMS") + num + "";

        }
        this.setState({
          LeaveID: LeaveID,

        });
        console.log(LeaveID);
      });

  }


  public UploadFile() {
    var reactHandler = this;

    var input: any = document.getElementById("leave-file-upload");
    var fileCount = input.files.length;
    console.log(fileCount);
    for (var i = 0; i < fileCount; i++) {
      var fileName = input.files[i].name;
      console.log(fileName);
      var file = input.files[i];
      var reader = new FileReader();

      reader.onload = (function (e) {
        return function (e) {
          //Push the converted file into array
          fileInfos.push({
            "name": file.name,
            "content": e.target.result
          });
        }
      })(file);
      reader.readAsArrayBuffer(file);
    }

    console.log("fileInfos:   " + fileInfos);

    reactHandler.setState({
      AttachmentCopies: fileInfos
    });

    console.log(this.state.AttachmentCopies);
  }
  public uploadfilestolist() {

    var reactHandler = this;
    var fileArr = [];


    const myfile = (document.querySelector(".leave-file-upload") as HTMLInputElement).files.length;
    for (var j = 0; j < myfile; j++) {
      let fileVal = (document.querySelector(".leave-file-upload") as HTMLInputElement).files[j];
      fileArr.push(fileVal);
      console.log("fileVal:  " + fileVal);
    }

    for (var i = 0; i < fileArr.length; i++) {

      var tempfilename = fileArr[i];

      const reader = new FileReader();
      reader.onload = (function (tempfilename) {


        return function (e) {
          fileInfos.push({
            "name": tempfilename.name,
            "content": e.target.result
          });
        }

      })(tempfilename);

      reader.readAsArrayBuffer(tempfilename);

    }

    console.log("fileInfos:   " + fileInfos);

    reactHandler.setState({
      AttachmentCopies: fileInfos
    });

    console.log(this.state.AttachmentCopies);
    // reactHandler.TriggerAttachment();

  }

  public checkIsSameOrBefore(date1: moment.MomentInput, date2: moment.MomentInput) {

    return moment(date1).isBefore(date2);

  }
  public checkIsSame(date1: moment.MomentInput, date2: moment.MomentInput) {

    return moment(date1).isSame(date2);
  }
  public checkIsafter(date1: moment.MomentInput, date2: moment.MomentInput) {

    return moment(date1).isAfter(date2);

  }
  public LeaveformValidation() {

    var Formstatus = false;
    var ErrorMsg = "";
    $("#divErrorText").empty();

    var Leavetype = $("#ddl-leavetype").val();
    var Startdate = $("#txt-Startdate").val();
    var EndDate = $("#txt-Enddate").val();
    var Day = $("#ddl-Day").val();
    var Time = $("#ddl-time").val();
    var Reason = $("#txt-reason").val();

    var start = moment(Startdate, "YYYY-MM-DD");
    var end = moment(EndDate, "YYYY-MM-DD");
    var Days = moment.duration(end.diff(start)).add(1, 'days').asDays();


    if (Leavetype == "") {

      ErrorMsg = "Please Select Leave Type";
      Formstatus = true;

    } else if (Formstatus == false && Startdate == "") {
      ErrorMsg = "Please Select StartDate";
      Formstatus = true;


    } else if (Formstatus == false && EndDate == "") {
      ErrorMsg = "Please Select EndDate";
      Formstatus = true;


    } else if (Formstatus == false && Day == "") {
      ErrorMsg = "Please Select Day";
      Formstatus = true;

    } else if (Formstatus == false && Reason == "") {
      ErrorMsg = "Please Enter Reason";
      Formstatus = true;

    }

    if (RestrictedHoliday == true) {
      if (this.state.dates.length == 0) {
        ErrorMsg = "Please Select Comp Off Date";
        Formstatus = true;
      } else if (Days != this.state.dates.length) {
        if (Days == 1) {
          ErrorMsg = "Please choose " + Days + " day to compensate for the restricted holiday";
        } else {
          ErrorMsg = "Please choose " + Days + " days to compensate for the restricted holiday";
        }
        Formstatus = true;
      }
    }


    if (Formstatus) {
      $("#divErrorText").append(ErrorMsg);
      $("#divErrorText").show();
      return false;
    } else {
      $("#divErrorText").empty();
      $("#divErrorText").hide();
      return true;
    }

    //End of for loop
  }
  public selectedleavetype(leavetype: string) {
    if (leavetype == 'Full Day') {
      $(".opttime").prop('disabled', true);
      $("#ddl-am-time").prop("checked", false);


    } else {

      $("#ddl-am-time").prop("checked", true);
      $(".opttime").prop('disabled', false);


      //To restrict future date

      // $('#txt-Enddate').attr('max', startdate.toString().slice(0, 10));
    }

  }


  public async DeleteAttachment(e: number) {


    swal({
      title: "Are you sure?",
      text: "You want to delete this attachment!",
      icon: "warning",
      buttons: ["No", "Yes"],
      dangerMode: true,
    }).then((willDelete) => {
      if (willDelete) {

        const temp = [this.state.AttachmentCopies];
        console.log(temp);
        // removing the element using splice
        temp.splice(e, 1);

        // updating the list
        this.setState({
          AttachmentCopies: temp
        });

        console.log(this.state.AttachmentCopies);


        swal({
          text: "File has been deleted",
          icon: "success",
          button: false,
          timer: 500
        } as any).then(() => {

        });
      }
    });
  }
  public handleFiles(f: { fileList: any; }) {

    var filelist = f.fileList;

    var fileInfos: IAttachmentFileInfo[] = [];

    // loop through files
    for (var i = 0; i < filelist.length; i++) {

      // get item
      let file: File = filelist.item(i);

      fileInfos.push({
        name: file.name,
        content: file
      });

    }
    this.setState({
      AttachmentCopies: fileInfos
    });


  }
  public handleDateChange(selectedDates: any) {
    const formattedDates = selectedDates.map((date: any) => date.format(format));
    console.log('Selected Dates:', formattedDates);

    this.setState({
      dates: formattedDates,
    });

  }
  public getEndDate() {
    var Date = $("#txt-Enddate").val()
    var CompensateMinDate = moment(Date, "YYYY-MM-DD").add(1, 'days').format("YYYY-MM-DD")
    this.setState({ dates: [] })
    if (Date == "") {
      this.setState({
        DatePickerDisable: true
      })
    } else {
      this.setState({
        DatePickerDisable: false,
        EndDate: CompensateMinDate,
      })
    }
  }
  public getStartDate() {
    var Date: any = $("#txt-Startdate").val()
    $('#txt-Enddate').attr('min', Date);
    $('#txt-Enddate').val("")
    this.setState({ dates: [] })
    if (Date == "") {
      this.setState({
        DatePickerDisable: true
      })
      $("#txt-Enddate").prop("disabled", true);
    } else {
      this.setState({
        DatePickerDisable: false,
      })
      $("#txt-Enddate").prop("disabled", false);
    }
  }

  public render(): React.ReactElement<ILeaveMgmtProps> {
    console.log(this.state.AttachmentCopies);
    let handler = this;

    const LeaveRequestAttachments: JSX.Element[] = this.state.AttachmentCopies.map(function (item, key) {
      var Extension = item.name.split(/\.(?=[^\.]+$)/);
      var Ext = Extension[1].toUpperCase(); //PDF   
      console.log(item.name);

      var Icon = "";
      if (Ext == "PDF") {
        Icon = `${handler.props.siteurl}/SiteAssets/LeavePortal/img/pdf.svg`;
      } else if (Ext == "JPG" || Ext == "JPEG" || Ext == "PNG" || Ext == "GIF" || Ext == "SVG") {
        Icon = `${handler.props.siteurl}/SiteAssets/LeavePortal/img/pdf.svg`;
      } else if (Ext == "DOCX" || Ext == "DOC") {
        Icon = `${handler.props.siteurl}/SiteAssets/LeavePortal/img/docx.svg`;
      } else if (Ext == "XLSX" || Ext == "XLS") {
        Icon = `${handler.props.siteurl}/SiteAssets/LeavePortal/img/xlsx.svg`;
      } else if (Ext == "PPTX") {
        Icon = `${handler.props.siteurl}/SiteAssets/LeavePortal/img/pptx.svg`;
      }
      return (
        <div className="file-img upload-img">
          <ul className="nav nav-pills">
            <li><img src={`${Icon}`} alt="image" className="attachment-img" /> </li>
            <li className="word-data"><p className="asset-info-header">{item.name}</p><p></p></li>
          </ul>
          <div className="close-doc-img"><a href="#" onClick={() => handler.DeleteAttachment(item.name)}><img src="https://tmxin.sharepoint.com/sites/ER/SiteAssets/LeavePortal/img/close (3).png" alt="close-icon" className="close-image" /></a></div>
        </div>
      );
    });

    return (
      <div className={styles.leaveMgmt} >
        <header>
          <div className="container">
            <div className="logo">
              <img src="https://tmxin.sharepoint.com/sites/ER/SiteAssets/LeavePortal/img/logo_small.png" alt="image" />
            </div>
            <div className="header-title"><h3>Leave Management System</h3></div>
            <div className="notification-part">
              <ul>
                <li className="person-details">
                  {/* <span id="CurrentUser-Profilepicture"> <img src={`${this.state.CurrentUserProfilePic}`} alt="image" /> <span>  </span>  </span>*/}
                  <span id="CurrentUser-displayname">{this.state.CurrentUserName}</span>

                  <a href="#" onClick={this.logout}><img src="https://tmxin.sharepoint.com/sites/ER/SiteAssets/LeavePortal/img/logout.png" /></a>
                </li>
              </ul>
            </div>
          </div>
        </header>

        <nav>
          <ul>
            <li className="active"><a href="https://tmxin.sharepoint.com/sites/ER/SitePages/Dashboard.aspx?env=WebView"> Home  </a> </li>
            <li> <a href="https://tmxin.sharepoint.com/sites/ER/SitePages/Aboutus.aspx?env=WebView"> About   </a> </li>
            <li> <a href="https://tmxin.sharepoint.com/sites/ER/SitePages/Holiday.aspx?env=WebView"> Holidays  </a> </li>
            <li> <a href="https://tmxin.sharepoint.com/sites/ER/SitePages/PermissionDashboard.aspx?env=WebView"> Permission  </a> </li>
          </ul>
        </nav>

        <div className="container">
          <div className="dashboard-wrap">

            <div className="form-header">
              <a href="https://tmxin.sharepoint.com/sites/ER/SitePages/Dashboard.aspx?env=WebView"><img src="https://tmxin.sharepoint.com/sites/ER/SiteAssets/LeavePortal/img/back.svg" alt="image" /></a><span>Leave Request</span>
            </div>

            <div className="form-body">
              <div className="form-section">
                <div className="row">
                  <div className="col-md-8">
                    <div className="form-group required relative">
                      <div className="form-group">
                        {/* <input type="text" className="form-control" placeholder='Employee Name' id="txt-name"  />*/}
                      </div>
                      {/* <span className="floating-label form-comment-txt "> </span>*/}
                    </div>
                  </div>
                </div>
                <div className="row">
                  <div className="col-md-4 col-sm-4">
                    <div className="form-group required relative">
                      <select name="leavetype" id="ddl-leavetype" className="form-control" onChange={() => this.clearerror()}>

                        <option value="">Select</option>
                        <option value="Earned Leave">Earned Leave</option>
                        <option value="Casual Leave">Casual Leave</option>
                        <option value="Sick Leave">Sick Leave</option>
                        <option value="Unpaid Leave">Unpaid Leave</option>
                        <option value="Maternity Leave">Maternity Leave</option>
                        <option value="Paternity Leave">Paternity Leave</option>
                        <option value="Restricted Leave">Restricted Leave</option>

                      </select>
                      <span className="floating-label" id="ddl-leave-type-label"> Leave Type </span>

                    </div>

                  </div>
                  <div className="col-md-4 col-sm-4">
                    <div className="form-group required relative  chk-box">

                      <span className="floating-label "> Day </span>

                      <label className="radio-inline">

                        <input type="radio" name="optradio" id="ddl-full-Day" value="Full Day" className="optday" onClick={() => this.selectedleavetype('Full Day')} />Full Day

                      </label>
                      <label className="radio-inline">
                        <input type="radio" name="optradio" id="ddl-half-Day" value="Half Day" className="optday" onClick={() => this.selectedleavetype('Half Day')} />Half Day

                      </label>

                    </div>
                  </div>
                  <div className="col-md-4 col-sm-4">
                    <div className="form-group required relative chk-box">
                      <span className="floating-label"> Time </span>

                      <label className="radio-inline">

                        <input type="radio" name="optradio1" id="ddl-am-time" value="First Half" className="opttime" />First Half

                      </label>
                      <label className="radio-inline">

                        <input type="radio" name="optradio1" id="ddl-pm-time" value="Second Half" className="opttime" />Second Half

                      </label>

                    </div>
                  </div>
                </div>
                <div className="row">
                  <div className="col-md-4 col-sm-4">
                    <div className="form-group required relative">
                      <div className="container">

                      </div>
                      {/*<input type="text" id="txt-Startdate"  className="form-control" placeholder="dd-mm-yyyy" autoComplete="off"/>*/}
                      <input type="date" className="form-control" id="txt-Startdate" autoComplete="off" onChange={() => this.getStartDate()} />
                      <span className="floating-label ">Start Date</span>

                    </div>
                  </div>

                  <div className="col-md-4 col-sm-4">
                    <div className="form-group required relative">
                      <input type="date" className="form-control" id="txt-Enddate" onChange={() => this.getEndDate()} />
                      {/* <input type="text" id="txt-Enddate" className="form-control"  placeholder="dd-mm-yyyy" autoComplete="off" />*/}
                      <span className="floating-label ">End Date</span>

                    </div>
                  </div>

                  <div className="col-md-4 col-sm-4" id='comp_off_date' style={{ display: "none" }}>
                    <div className="form-group required relative">
                      <div style={{ textAlign: "center" }}>
                        <DatePicker
                          value={this.state.dates}
                          onChange={(selectedDates) => this.handleDateChange(selectedDates)}
                          multiple
                          sort
                          format={format}
                          calendarPosition="bottom-center"
                          plugins={[<DatePanel />]}
                          placeholder="Select dates"
                          minDate={this.state.EndDate}
                          disabled={this.state.DatePickerDisable}
                        />
                      </div>
                      <span className="floating-label ">Compensation Day Selection</span>

                    </div>
                  </div>

                </div>

                <div className="row">
                  <div className="col-md-8">
                    <div className="form-group required relative">
                      <div className="form-group">

                        {/*<input type="text" className="form-control" id="txt-reason" maxLength={250} autoComplete="off" onKeyPress={() => this.LeaveformValidation()} />*/}
                        <input type="text" className="form-control" id="txt-reason" maxLength={250} autoComplete="off" onKeyPress={() => this.clearerror()} />
                        <span className="floating-label ">Enter Reason</span>
                      </div>

                    </div>
                  </div>
                </div>

                <div className="row">
                  <div className="col-md-12">
                    <div className="image-upload">

                      <ReactFileReader id="leave-file-upload" className="leave-file-upload" multipleFiles={false} fileTypes={[".csv", ".xlsx", ".Docx", ".pdf", ".png", ".jpeg", ".jpg", ".svg"]} base64={true}
                        handleFiles={(f: any) => this.handleFiles(f)}  >
                        <label htmlFor="leave-file-upload" className="img-upload">
                          <img src="https://tmxin.sharepoint.com/sites/ER/SiteAssets/LeavePortal/img/upload.png" className="upload_file" />
                          <h5>Choose an file. </h5>

                        </label>
                        {/* <button className="img-upload">Choose an file. </button>*/}
                      </ReactFileReader>
                      {/*  <input id="leave-file-upload" className="leave-file-upload" name="leave-file-upload" type="file" onChange={() => this.UploadFile} multiple />*/}
                    </div>
                    <div id="leaveBindCopy">
                      {LeaveRequestAttachments}
                      {/*  {handler.state.AttachmentCopies && handler.state.AttachmentCopies.map(function (item, key) {
                        var Extension = item.name.split(/\.(?=[^\.]+$)/);
                        console.log(item.name);
                        var Ext = Extension[1].toUpperCase(); //PDF         
                        var Icon = "";
                        if (Ext == "PDF") {
                          //  Icon = `${this.props.siteurl}/SiteAssets/LeavePortal/img/pdf.svg`;
                          Icon = `${handler.props.siteurl}/SiteAssets/LeavePortal/img/pdf.svg`;
                        } else if (Ext == "JPG" || Ext == "JPEG" || Ext == "PNG" || Ext == "GIF" || Ext == "SVG") {
                          Icon = `${handler.props.siteurl}/SiteAssets/LeavePortal/img/DummyImage.svg`;
                        } else if (Ext == "DOCX" || Ext == "DOC") {
                          Icon = `${handler.props.siteurl}/SiteAssets/LeavePortal/img/docx.svg`;
                        } else if (Ext == "XLSX" || Ext == "XLS") {
                          Icon = `${handler.props.siteurl}/SiteAssets/LeavePortal/img/xlsx.svg`;
                        } else if (Ext == "PPTX") {
                          Icon = `${handler.props.siteurl}/SiteAssets/LeavePortal/img/pptx.svg`;
                        }
                        console.log(item.name);
                        return (
                          <div className="file-img upload-img">
                            <ul className="nav nav-pills">
                              <li><img src={`${Icon}`} alt="image" className="attachment-img" /> </li>
                              <li className="word-data"><p className="asset-info-header">{item.name}</p><p></p></li>

                            </ul>
                            <div className="close-doc-img"><a href="#"><img src="https://tmxin.sharepoint.com/sites/ER/SiteAssets/LeavePortal/img/close (3).png" alt="close-icon" className="close-image" /></a></div>
                          </div>
                        );
                      })}*/}
                    </div>

                  </div>

                </div>

              </div>
              <div className="row">
                <div
                  className="alert alert-danger"
                  role="alert"
                  id="divErrorText"
                  style={{ display: "none" }}
                ></div>
                <div className="col-md-12 btn-padding">
                  <button className="btn btn-primary" id="submit" onClick={() => this.leavetypevalidation()}>Submit</button>
                  {/* <button className="btn btn-primary" id="submit" onClick={() => this.LeaveformValidation()}>Submit</button> */}

                </div>
              </div>
            </div>
          </div>
        </div>
      </div>



    );
  }
  public componentDidUpdate(prevProps: Readonly<ILeaveMgmtProps>, prevState: Readonly<ILeaveMgmtState>, snapshot?: any): void {


  }
}
