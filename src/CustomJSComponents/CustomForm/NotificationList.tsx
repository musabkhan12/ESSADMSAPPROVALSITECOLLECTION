import React from 'react';
import moment from 'moment';

import { X,Settings } from 'react-feather';
 
interface Notification {
    Id: string;
    Created: string; // or Date, depending on your data type
    ContentName: string;
    ContentComment: string;
 
    ActionUser?: { Title: string };
    ContentType0: string;
    NotifiedUser?: { Title: string };
}
 
interface CategorizedNotifications {
    today: Notification[];
    yesterday: Notification[];
    earlier: Notification[];
}
 
const NotificationList = ({ NotificationArray, handleNotificationClick, OnClearall }: any) => {
    // Function to categorize notifications
   
 
    const categorizeNotifications = (notifications: Notification[]) => {
        const now = moment();
        const categorized: CategorizedNotifications = { today: [], yesterday: [], earlier: [] };
 
        notifications.forEach((notify) => {
            const createdDate = moment(notify.Created);
 
            if (createdDate.isSame(now, 'day')) {
                categorized.today.push(notify);
            } else if (createdDate.isSame(moment().subtract(1, 'day'), 'day')) {
                categorized.yesterday.push(notify);
            } else {
                categorized.earlier.push(notify);
            }
        });
 
        return categorized;
    };
 
    // Categorize notifications
    const categorizedNotifications = categorizeNotifications(NotificationArray);
    console.log(categorizedNotifications.yesterday, 'categorizedNotifications.yesterday');
    const goToNext = () => {
   
        let webUrl = window.location.href;
        const baseUrl = webUrl.substring(0, webUrl.lastIndexOf("/SitePages") + "/SitePages".length);
        console.log(baseUrl); // Output: https://example.com/webUrl/sites/AlRostmani/SitePages
        window.location.href = `${baseUrl}/NotificationDetails.aspx`;
    }
    const goToSettings = () => {
   
        let webUrl = window.location.href;
        const baseUrl = webUrl.substring(0, webUrl.lastIndexOf("/SitePages") + "/SitePages".length);
        console.log(baseUrl); // Output: https://example.com/webUrl/sites/AlRostmani/SitePages
        window.location.href = `${baseUrl}/ManageNotification.aspx`;
    }
    return (
        <div>
            <div className="flex">
                <div className="row">
                    <div className="col-md-8">
                        <h5 className="p-1 font-16 text-dark">Notifications</h5>
                    </div>
                    {/* <div style={{textAlign:'right'}} className="col-md-4">
                    <h5 style={{ textDecoration: 'underline', cursor: 'pointer' }} className="pt-2 font-12" onClick={() => OnClearall('Clear')} >Clear All</h5>
                    </div> */}
                </div>
            </div>
  <div className="heightlist">
            {NotificationArray.length > 0 ? (
                <div>
                    {/* Today Section */}
                    {categorizedNotifications.today.length > 0 && (
                        <div>
                            <h6 className="text-dark font-14 fw-bold p-2 pb-0">Today</h6>
                            {categorizedNotifications.today.map((notify) => (
                               
                                <a key={notify.Id} className="dropdown-item p-2 notify-item card unread-noti shadow-none mb-1"
                                    onClick={() => handleNotificationClick(notify)} style={{display: 'flex', margin: '0 auto' }}>
                                    <div className="">
                                        <span className="float-start noti-close-btn text-muted">
                                            <img  src={require("../../CustomAsset/Noti.jpg")}
                                       
                                        className="me-1 noti"
                                      /></span>
                                        <div className="d-flex align-items-center">
                                            <div className="flex-grow-1 text-truncate ms-2">
                                                <h5 className="noti-item-title hovertext  fw-semibold font-14 mb-0" style={{ textTransform: 'capitalize' }}>
                                                {notify.ContentName || notify.ContentComment}
                                                    <small className="fw-normal text-muted ms-0" style={{ textTransform: 'lowercase' }}>{moment(notify.Created).fromNow()}</small>
                                                </h5>
                                                <small className="noti-item-subtitle text-muted">{notify?.ActionUser?.Title} {notify.ContentType0} on {notify?.NotifiedUser?.Title}</small>
                                            </div>
                                            <div className="flex-shrink-0">
                                                {/* <div className="notify-icon bg-primary">
                                                    <X style={{ background: '#fff' }} />
                                                </div> */}
                                            </div>
                                        </div>
                                    </div>
                                </a>
                            ))}
                        </div>
                    )}
 
                    {/* Yesterday Section */}
                    {categorizedNotifications.yesterday.length > 0 && (
                        <div>
                            <h6 className="text-dark font-14 fw-bold p-2 pb-0">Yesterday</h6>
                            {categorizedNotifications.yesterday.map((notify) => (
                                <a key={notify.Id} className="dropdown-item p-2 notify-item card unread-noti shadow-none mb-1"
                                    onClick={() => handleNotificationClick(notify)} style={{ display: 'flex', margin: '0px auto' }}>
                                    <div className="">
                                    <span className="float-start noti-close-btn text-muted">
                                            <img  src={require("../../CustomAsset/Noti.jpg")}
                                       
                                        className="me-1 noti"
                                      /></span>
                                        <div className="d-flex align-items-center">
                                            <div className="flex-grow-1 text-truncate ms-2">
                                                <h5 className="noti-item-title hovertext  fw-semibold font-14 mb-0" style={{ textTransform: 'capitalize' }}>
                                                {notify.ContentName || notify.ContentComment}
                                                    <small className="fw-normal text-muted ms-0">{moment(notify.Created).fromNow()}</small>
                                                </h5>
                                                <small className="noti-item-subtitle text-muted">{notify?.ActionUser?.Title} {notify.ContentType0} on {notify?.NotifiedUser?.Title}</small>
                                            </div>
                                            <div className="flex-shrink-0">
                                                {/* <div className="notify-icon bg-primary">
                                                    <X style={{ background: '#fff' }} />
                                                </div> */}
                                            </div>
                                        </div>
                                    </div>
                                </a>
                            ))}
                        </div>
                    )}
 
                    {/* Earlier Section */}
                    {categorizedNotifications.earlier.length > 0 && (
                        <div>
                            <h6 className="text-dark font-14 fw-bold p-2 pb-0">Earlier</h6>
                            {categorizedNotifications.earlier.map((notify) => (
                                <a key={notify.Id} className="dropdown-item p-2 notify-item card unread-noti shadow-none mb-1"
                                    onClick={() => handleNotificationClick(notify)} style={{ display: 'flex', margin: '0px auto' }}>
                                    <div className="">
                                    <span className="float-start noti-close-btn text-muted">
                                            <img  src={require("../../CustomAsset/Noti.jpg")}
                                       
               className="me-1 noti"
                                      /></span>
                                        <div className="d-flex align-items-center">
                                            <div className="flex-grow-1 text-truncate ms-2">
                                                <h5 className="noti-item-title hovertext  fw-semibold font-14 mb-0" style={{ textTransform: 'capitalize' }}>
                                                {notify.ContentName || notify.ContentComment} <small className="fw-normal text-muted ms-0">{moment(notify.Created).fromNow()}</small>
                                                </h5>
                                                <small className="noti-item-subtitle text-muted">{notify?.ActionUser?.Title} {notify.ContentType0} on {notify?.NotifiedUser?.Title}</small>
                                            </div>
                                            <div className="flex-shrink-0">
                                                {/* <div className="notify-icon bg-primary">
                                                    <X style={{ background: '#fff' }} />
                                                </div> */}
                                            </div>
                                        </div>
                                    </div>
                                </a>
                            ))}
                        </div>
                    )}
                </div>
            ) : (
                <p>No notifications</p>
            )}
        </div>
        <h5 style={{textDecoration:'underline',float:'left'}} className="p-3 font-12 hovertext mb-0 text-center" onClick={goToNext}>View All</h5>
        <h5 style={{float:'right'}} className="p-3 font-12 text-center mb-0" onClick={goToSettings}> <Settings size={18} /> </h5>
        </div>
    );
};
 
export default NotificationList;