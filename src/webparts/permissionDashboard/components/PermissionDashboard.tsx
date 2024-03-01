import * as React from 'react';
import styles from './PermissionDashboard.module.scss';
import { IPermissionDashboardProps } from './IPermissionDashboardProps';
import { IPermissionDashboardState } from './IPermissionDashboardState';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from "@microsoft/sp-loader";
import { sp } from "@pnp/sp";
import { Web } from '@pnp/sp/webs';

import "datatables.net-dt/js/dataTables.dataTables";
import "datatables.net-dt/css/jquery.dataTables.min.css";
//import "datatables.net-dt/js/dataTables.dataTables";
//import "datatables.net-dt/css/jquery.dataTables.min.css";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import * as $ from 'jquery';
import swal from "sweetalert";

import * as moment from 'moment';
let ItemId;
var CurrentUSERNAME = "";
var Usertype = "";
const NewWeb = Web('https://tmxin.sharepoint.com/sites/ER/');
export default class PermissionDashboard extends React.Component<IPermissionDashboardProps, IPermissionDashboardState> {

  public constructor(props: IPermissionDashboardProps) {

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

    SPComponentLoader.loadScript(
      `https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js`
    );

    SPComponentLoader.loadScript(
      `https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.js`
    );
    SPComponentLoader.loadCss(
      `https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css`
    );

    SPComponentLoader.loadCss(
      `https://tmxin.sharepoint.com/sites/ER/SiteAssets/LeavePortal/css/style.css?v=1.14`
    );


    super(props);
    sp.setup({
      spfxContext: this.props.context
    });

    this.state = {
      DatatableItems: [],
      Loggedinuserid: null,
      IsAdmin: false,
      CurrentUserName: "",
      CurrentUserDesignation: "",
      CurrentUserProfilePic: "",
      CurrentUserId: null,
      IsUser: false,
      Empemail: ""
    };
  }
  // public async isOwnerGroupMember() {
  //   var reacthandler = this;
  //   let userDetails = await this.spLoggedInUser(this.props.context);

  //   let userID = userDetails.Id;
  //   console.log(userID);
  //   $.ajax({

  //     // url: `${reacthandler.props.siteurl}/_api/web/sitegroups/getByName('LMS Admin')/Users?$filter=Id eq  + ${this.props.userId}`,
  //     url: `${reacthandler.props.siteurl}/_api/web/sitegroups/getByName('LMS Admin')/Users?$filter=Id eq ${userID}`,

  //     type: "GET",

  //     headers: { 'Accept': 'application/json; odata=verbose;' },

  //     success: function (resultData) {

  //       if (resultData.d.results.length == 0) {
  //         console.log("User not in group : LMS Admin Owners");
  //         setTimeout(() => {
  //           reacthandler.GetUserlistitems();
  //         }, 1000);

  //       } else {
  //         console.log("User in group : LMS Admin Owners");
  //         setTimeout(() => {
  //           reacthandler.GetAdminlistitems();
  //         }, 1000);
  //       }

  //     },

  //     error: function (jqXHR, textStatus, errorThrown) {
  //       console.log("Error while checking user in Owner's group");
  //     }

  //   });

  // }
  public async isOwnerGroupMember() {
    var reacthandler = this;
    let userDetails = await this.spLoggedInUser(this.props.context);

    let userID = userDetails.Id;
    console.log(userID);
    $.ajax({

      // url: `${reacthandler.props.siteurl}/_api/web/sitegroups/getByName('LMS Admin')/Users?$filter=Id eq  + ${this.props.userId}`,
      url: `https://tmxin.sharepoint.com/sites/lms/_api/web/sitegroups/getByName('LMS Admin')/Users?$filter=Id eq ${userID}`,

      type: "GET",

      headers: { 'Accept': 'application/json; odata=verbose;' },

      success: function (resultData) {

        if (resultData.d.results.length == 0) {
          console.log("User not in group : LMS Admin Owners");
          setTimeout(() => {
            reacthandler.GetUserlistitems();
          }, 1000);


        } else {

          console.log("User in group : LMS Admin Owners");

          setTimeout(() => {

            reacthandler.GetAdminlistitems();
          }, 1000);

        }

      },

      error: function (jqXHR, textStatus, errorThrown) {
        console.log("Error while checking user in Owner's group");
      }

    });

  }
  public GetCurrentUserDetails() {

    var reacthandler = this;

    $.ajax({

      url: `${reacthandler.props.siteurl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`,

      type: "GET",

      headers: { 'Accept': 'application/json; odata=verbose;' },

      success: function (resultData) {

        var email = resultData.d.Email;

        var Name = resultData.d.DisplayName;

        var Designation = resultData.d.Title;

        reacthandler.setState({

          CurrentUserName: Name,

          CurrentUserDesignation: Designation,
          Empemail: email,

          CurrentUserProfilePic: `${reacthandler.props.siteurl}/_layouts/15/userphoto.aspx?size=l&username=${email}`

        });

      },

      error: function (jqXHR, textStatus, errorThrown) {

      }

    });

  }

  public GetUserlistitems() {
    var reactHandler = this;

    NewWeb.lists.getByTitle("EmployeePermission").items.select("Id", "PermissionHour", "TimeUpto", "PermissionOn", "timefromwhen", "Requester", "EmployeeEmail", "Reason", "Status").filter(`Author/Id eq ${this.props.userId}`).orderBy("Created", false).top(5000).get()
      //NewWeb.lists.getByTitle("EmployeePermission").items.select("Id", "PermissionHour", "TimeUpto", "PermissionOn", "timefromwhen", "Requester","EmployeeEmail", "Reason", "Status").filter("EmployeeEmail eq '" + this.state.Empemail + "'").orderBy("Created", false).top(5000).get()

      .then((items) => {
        if (items.length != 0) {

          reactHandler.setState({
            DatatableItems: items
          });
          this.loadTable();
          $("#no_data").hide()

        }
        else {
          this.loadTable();
          $("#no_data").show()
        }
      });


  }
  public GetAdminlistitems() {
    this.setState({ IsAdmin: true });
    var reactHandler = this;

    NewWeb.lists.getByTitle("EmployeePermission").items.select("Id", "PermissionHour", "TimeUpto", "PermissionOn", "timefromwhen", "Requester", "EmployeeEmail", "Reason", "Status").orderBy("Created", false).top(5000).get()
      //NewWeb.lists.getByTitle("EmployeePermission").items.select("Id", "PermissionHour", "TimeUpto", "PermissionOn", "timefromwhen", "Requester","EmployeeEmail", "Reason", "Status").filter("EmployeeEmail eq '" + this.state.Empemail + "'").orderBy("Created", false).top(5000).get()

      .then((items) => {
        if (items.length != 0) {

          reactHandler.setState({
            DatatableItems: items
          });
          this.loadTable();

        }
        else {
          this.loadTable();
        }
      });


  }
  public async Checkusertype(UserType: string) {
    var reactHandler = this;
    if (UserType == "User") {

      await NewWeb.lists.getByTitle("EmployeePermission").items.select("Id", "PermissionHour", "TimeUpto", "PermissionOn", "timefromwhen", "Requester", "EmployeeEmail", "Reason", "Status").filter("EmployeeEmail eq '" + this.state.Empemail + "'").orderBy("Created", false).top(5000).get()

        .then((items) => {
          if (items.length != 0) {

            reactHandler.setState({
              DatatableItems: items
            });
            this.loadTable();

          }
          else {
            this.loadTable();
          }
        });

    } else {
      //if (UserType =="Admin") {
      await NewWeb.lists.getByTitle("EmployeePermission").items.select("Id", "PermissionHour", "TimeUpto", "PermissionOn", "timefromwhen", "Requester", "EmployeeEmail", "Reason", "Status").orderBy("Created", false).top(5000).get()

        .then((items) => {
          if (items.length != 0) {

            reactHandler.setState({

              DatatableItems: items
            });
            this.loadTable();

          }
          else {
            this.loadTable();
          }
        });

    }
  }
  public logout() {

    localStorage.clear();
    window.location.href = `https://login.windows.net/common/oauth2/logout`;

  }
  public _spLoggedInUserDetails() {
    NewWeb.currentUser.get().then((user) => {
      let userID = user.Id;
      this.setState({ CurrentUserId: userID });
    }, (errorResponse) => {
      //console.log(errorResponse);
    }
    );
  }
  private async spLoggedInUser(ctx: any) {
    try {
      const web = Web(ctx.pageContext.site.absoluteUrl);
      return await web.currentUser.get();
    } catch (error) {
      console.log("Error in spLoggedInUserDetails : " + error);
    }
  }
  public async componentDidMount() {
    this.GetCurrentUserDetails();
    const url: any = new URL(window.location.href);
    url.searchParams.get("ItemID");
    ItemId = url.searchParams.get("ItemID");

    let userDetails = await this.spLoggedInUser(this.props.context);
    console.log(userDetails.Id);
    let userID = userDetails.Id;
    this.setState({ CurrentUserId: userID });
    await this.isOwnerGroupMember();
    //await this.GetListitems();
    // console.log("User Type:"+Usertype);
    // this.Checkusertype(Usertype);


    //  this.loadTable();


  }
  public GetPermissionDetails() {
    var reactHandler = this;
    //  AnnualArr = [];
    //SickArr = [];

    var url = "https://tmxin.sharepoint.com/sites/ER/_api/web/lists/getbytitle('EmployeePermission')/items?$select=PermissionHour,TimeUpto,PermissionOn,timefromwhen,Requester,Reason,Status&$orderby=Created desc";

    $.ajax({
      url: url,
      type: "GET",
      async: false,
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        console.log(resultData);
        reactHandler.setState({

          DatatableItems: resultData.d.results
        });


        setTimeout(() => {
          reactHandler.loadTable();
        }, 1000);

        {/* for (var i = 0; i < resultData.d.results.length; i++) {
          if (resultData.d.results[i].LeaveType == "AnualLeave") {
            AnnualArr.push(resultData.d.results[i]);


          }
        }

        var TotalAnualLeave = `${AnnualArr.length}/${ttlAnnulLeave}`
      $("#annualLeave").html(TotalAnualLeave)*/}
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });
  }
  public async GetListitems() {
    var reactHandler = this;
    var UserType = "";
    let groups = await NewWeb.currentUser.groups();
    for (var i = 0; i < groups.length; i++) {
      if (groups[i].Title == 'LMS Admin') {
        UserType = "Admin";
        // this.setState({ IsAdmin: true });

        console.log(UserType);
        Usertype = UserType;
        return Usertype;

      } else {

        UserType = "User";
        Usertype = "User"

      }
    }

    return Usertype;
    {/*if (UserType =="User") {

      await NewWeb.lists.getByTitle("EmployeePermission").items.select("Id", "PermissionHour", "TimeUpto", "PermissionOn", "timefromwhen", "Requester", "Reason", "Status").filter(`Author/Id eq ${this.state.CurrentUserId}`).orderBy("Created", false).top(5000).get()
        .then((items) => {
          if (items.length != 0) {

            reactHandler.setState({
              DatatableItems: items
            });
            this.loadTable();

          }
          else{
            this.loadTable();
          }
     
        });

    } else {
     //if (UserType =="Admin") {

      await NewWeb.lists.getByTitle("EmployeePermission").items.select("Id", "PermissionHour", "TimeUpto", "PermissionOn", "timefromwhen", "Requester", "Reason", "Status").orderBy("Created", false).top(5000).get()

        .then((items) => {
          if (items.length != 0) {

            reactHandler.setState({

              DatatableItems: items
            });
            this.loadTable();

          }
          else{
            this.loadTable();
          }
     
        });

    }*/}

  }


  public Displaypermissionform() {

    location.href = `https://tmxin.sharepoint.com/sites/ER/SitePages/Permission.aspx?env=WebView`;
  }

  public GetCurrentUserName() {
    var reactHandler = this;
    $.ajax({
      url: `${this.props.siteurl}/_api/web/currentUser`,
      method: "GET",
      headers: {
        Accept: "application/json; odata=verbose",
      },
      success: function (data) {
        CurrentUSERNAME = data.d.Title;
        reactHandler.setState({
          CurrentUserName: CurrentUSERNAME

        });
      },
      error: function (data) { },
    });
  }
  public loadTable() {
    //($('#LMSDashboard') as any).DataTable.destroy();

    ($('#LMSDashboard') as any).DataTable({
      pageLength: 5,
      "bSort": false,
      "bDestroy": true,

      lengthMenu: [[5, 10, 20, 50, -1], [5, 10, 20, 50, "All"]],

      initComplete: function () {

        this.api().columns().every(function () {

          var column = this;
          //  var select = $('<select class="form-control"><option value="">All</option></select>')
          var select = $('<select><option value="">All</option></select>')

            .appendTo($(column.header()).empty()).on('change', function () {

              var val = ($ as any).fn.dataTable.util.escapeRegex(

                ($(this) as any).val()

              );

              column.search(val ? '^' + val + '$' : '', true, false).draw();


            });

          column.data().unique().sort().each(function (d: string, j: any) {

            var temp2 = d;
            if (temp2.indexOf(">") != -1) {
              var temp = d.split(">");
              var temporary = temp[3].split("<");
              select.append('<option value="' + temporary[0] + '">' + temporary[0] + '</option>')
            } else {
              select.append('<option value="' + d + '">' + d + '</option>')
            }

          });


        });

      }

    });
    //  }, 500);

  }
  public Cancel_Request_(itemidno: number) {

    swal({
      title: ` "Are you sure?"`,
      text: "Would you like to cancel the permission?",
      icon: "warning",
      buttons: ["No", "Yes"],
      dangerMode: true,
    } as any).then((willdelete) => {
      if (willdelete) {
        NewWeb.lists.getByTitle("EmployeePermission").items.getById(itemidno).update({
          Status: "Cancelled",
          CancelledBy: this.state.CurrentUserName
        }).then(() => {
          swal({
            text: "Permission cancelled successfully!",
            icon: "success",
          }).then(() => {
            location.reload()
          });

        })



      }
    })
  }
  public render(): React.ReactElement<IPermissionDashboardProps> {
    let count = 0;
    let handler = this;

    const DataTableBodycontent: JSX.Element[] = this.state.DatatableItems.map(function (item, key) {

      count++;
      return (

        <tr id={`${key}-row-id`}>
          <td>{key + 1}</td>
          {handler.state.IsAdmin == true &&
            <td>{item.Requester}</td>
          }
          <td>{moment(item.PermissionOn).format('DD-MM-YYYY')}</td>
          <td>{item.timefromwhen}</td>
          <td>{item.TimeUpto}</td>
          <td>{item.PermissionHour}</td>
          <td className="reason-td">{item.Reason}</td>

          {item.Status == "Pending" ?

            <td className="status pending text-center">{item.Status}</td>
            :
            item.Status == "Approved" ?
              <td className="status approved text-center">{item.Status}</td>
              :
              item.Status == "Rejected" ?
                <td className="status rejected text-center">{item.Status}</td>
                :
                (item.Status == "Cancelled" || item.Status == "Cancel") ?
                  <td className="status rejected text-center">{item.Status}</td>
                  :
                  <></>
          }
          <td>
            {(item.State !== "Cancel" || item.State !== "Cancelled") &&
              <>
                {(handler.state.Empemail == item.EmployeeEmail && item.Status != "Cancelled" && item.Status != "Rejected" && moment(item.timefromwhen, "DD-MM-YYYY hh:mm A").isSameOrAfter(moment(), 'day')) &&


                  <p onClick={() => handler.Cancel_Request_(item.Id)}><img src="https://tmxin.sharepoint.com/sites/ER/SiteAssets/LeavePortal/img/cancel.svg" alt="image" /></p>
                }

              </>
            }
          </td>
          {/* <td>{moment(item.TimeUpto, "YYYY-MM-DDTHH:mm").format('DD-MM-YYYY hh:mm A')}</td>   
          <td>{moment(item.timefromwhen,"YYYY-MM-DDTHH:mm").format('DD-MM-YYYY hh:mm A')}</td>
          <td><a href="#" onClick={() => handler.View(item.Id)}>View</a></td>*/}
        </tr>

      );
    });

    return (
      <div className={styles.permissionDashboard} >
        <header>
          <div className="container">
            <div className="logo">
              <img src="https://tmxin.sharepoint.com/sites/ER/SiteAssets/LeavePortal/img/logo_small.png" alt="image" />
            </div>
            <div className="header-title"><h3>Leave Management System</h3></div>
            <div className="notification-part">
              <ul>
                <li className="person-details">
                  {/*<span id="CurrentUser-Profilepicture"> <img src={`${this.state.CurrentUserProfilePic}`} alt="image" /> <span>  </span>  </span>*/}
                  <span id="CurrentUser-displayname">{this.state.CurrentUserName}</span>
                  <a href="#" onClick={this.logout}><img src="https://tmxin.sharepoint.com/sites/ER/SiteAssets/LeavePortal/img/logout.png" /></a>
                </li>
              </ul>
            </div>
          </div>
        </header>
        <nav>
          <ul>
            <li><a href="https://tmxin.sharepoint.com/sites/ER/SitePages/Dashboard.aspx?env=WebView"> Home  </a> </li>
            <li> <a href="https://tmxin.sharepoint.com/sites/ER/SitePages/Aboutus.aspx?env=WebView"> About   </a> </li>
            <li> <a href="https://tmxin.sharepoint.com/sites/ER/SitePages/Holiday.aspx?env=WebView"> Holidays  </a> </li>
            <li className="active"> <a href="https://tmxin.sharepoint.com/sites/ER/SitePages/PermissionDashboard.aspx?env=WebView"> Permission  </a> </li>
          </ul>
        </nav>
        <div className="container">
          <div className="dashboard-wrap">

            <div className="tab-headings clearfix">
              <ul className="nav nav-pills">


              </ul>

              {this.state.IsAdmin == true &&
                <a href="https://tmxin.sharepoint.com/sites/ER/Lists/EmployeePermission/Approvedlist.aspx" className="btn btn-outline leave-req-link " id="submit">View permission list</a>}
              <button className="btn btn-outline" id="submit" onClick={() => this.Displaypermissionform()}> Permission Request  </button>
            </div>

            <div className="tab-content">
              <div id="home" className="tab-pane fade in active">

                <div className="table-wrap">
                  <div className="table-search-wrap clearfix">
                    <div className="table-search relative">
                      {/* <input type="text" placeholder="Search Here" className="" />
                      <img src="https://tmxin.sharepoint.com/sites/POC/SPIP/SiteAssets/LeavePortal/img/search.svg" alt="image" />*/}
                    </div>

                  </div>
                  <table id="LMSDashboard" className="table" >
                    <thead>
                      <tr>
                        <th></th>
                        {this.state.IsAdmin == true &&
                          <th></th>
                        }
                        <th></th>
                        <th></th>
                        <th></th>
                        <th></th>
                        <th className="reason-select-input"></th>
                        <th className="text-center"> Status  </th>
                        <th></th>
                      </tr>
                    </thead>
                    <thead>

                      <tr>

                        <th>S.No</th>
                        {this.state.IsAdmin == true &&
                          <th>Employee Name</th>
                        }
                        <th>Requested On</th>
                        <th>Start Time</th>
                        <th>End Time</th>
                        <th>Permission Hours</th>
                        <th className="reason-td">Reason</th>
                        <th className="text-center"> Status  </th>
                        <th>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {DataTableBodycontent}
                    </tbody>

                  </table>
                </div>

              </div>

            </div>
          </div>
        </div>
      </div>
    );
  }
}
