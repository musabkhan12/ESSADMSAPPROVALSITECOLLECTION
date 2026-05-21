import React, { useEffect, useState, useRef } from 'react'
import g1 from "../../CustomAsset/Officials.jpg";
import "bootstrap/dist/css/bootstrap.min.css";
import "../CustomNewsWebpartTemplate/CustomNewsWebpartTemplate.scss";
import "../../CustomCss/mainCustom.scss";
import "../../Assets/Figtree/Figtree-VariableFont_wght.ttf";
import { Share2 } from 'feather-icons-react';
import { Bookmark } from 'feather-icons-react';
import { Calendar, Link, Share } from 'feather-icons-react';
import { addActivityLeaderboard } from "../../APISearvice/CustomService";
import { getAnncouncement } from "../../APISearvice/AnnouncementsService";
import moment from 'moment';
import { encryptId } from "../../APISearvice/CryptoService";
const CustomWebpartTemplate = ({ _sp, SiteUrl }) => {
    const [copySuccess, setCopySuccess] = useState('');
    const [show, setShow] = useState(false)
    const [AnnouncementData, setAnnouncement] = useState([])
    const [showDropdownId, setShowDropdownId] = React.useState(null);
    const [currentEmail, setEmail] = React.useState('');
    const [isMenuOpenshare, setIsMenuOpenshare] = useState(false);
    const menuRef = useRef(null);
    useEffect(() => {
        ApIcall()
        const handleClickOutside = (event) => {
            if (menuRef.current && !menuRef.current.contains(event.target)) {
                setIsMenuOpenshare(false);
            }
        };

        document.addEventListener('mousedown', handleClickOutside);

        return () => {
            document.removeEventListener('mousedown', handleClickOutside);
        };
    }, [_sp, SiteUrl])

    const ApIcall = async () => {
        setAnnouncement(await getAnncouncement(_sp))

    }
    const truncateText
        = (text, maxLength) => {
            if (!text)
                return "";
            // If text is null or undefined, return an empty string
            return text.length > maxLength ? text.
                substring
                (0, maxLength) +
                "..."
                : text;
        };

    const gotoAnnouncementDetails = async (valurArr) => {
        localStorage.setItem("AnnouncementId", valurArr.Id)
        localStorage.setItem("announcementArr", JSON.stringify(valurArr))
        await addActivityLeaderboard(_sp, "Announcement Click");
        setTimeout(() => {

            //   let IdStr=  encryptId(String(valurArr.Id))
            window.location.href = `${SiteUrl}/SitePages/AnnouncementDetails.aspx?${valurArr.Id}`;
            // window.location.href = `${SiteUrl}/SitePages/AnnouncementDetails.aspx`;

        }, 1000);
    }
    const copyToClipboard = (Id) => {
        const link = `${SiteUrl}/SitePages/AnnouncementDetails.aspx?${Id}`;
        navigator.clipboard.writeText(link)
            .then(() => {
                setCopySuccess('Link copied!');
                setTimeout(() => setCopySuccess(''), 2000); // Clear message after 2 seconds
            })
            .catch(err => {
                setCopySuccess('Failed to copy link');
            });
    };



    const toggleDropdown = (itemId) => {
        if (showDropdownId === itemId) {
            setShowDropdownId(null); // Close the dropdown if already open
        } else {
            setShowDropdownId(itemId); // Open the dropdown for the clicked item
        }
    };

    const sendanEmail = (item) => {
        // window.open("https://outlook.office.com/mail/inbox");

        // const subject = "Announcement Title-" + item.Title;
        // const body = 'Here is the link to the Announcement:' + `${SiteUrl}/SitePages/AnnouncementDetails.aspx?${item.Id}`;

        // const mailtoLink = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

        // // Open the link to launch the default mail client (like Outlook)
        // //window.location.href = mailtoLink;
        // const office365MailLink = `https://outlook.office.com/mail/deeplink/compose?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
        const subject = "Thought You’d Find This Interesting!";
        const body = 'Hi,' +
            'I came across something that might interest you: ' +
            `<a href="${siteUrl}/SitePages/AnnouncementDetails.aspx?${item.Id}"></a>`
        'Let me know what you think!';

        // window.open(`https://outlook.office365.com/mail/deeplink/compose?body= Here is the group link: "${siteUrl}/SitePages/GroupandTeamDetails.aspx?${idNum}"`);
        const office365MailLink = `https://outlook.office.com/mail/deeplink/compose?subject=${subject}&body=${body}`;

        window.open(office365MailLink, '_blank');
    };
    const [visibleItems, setVisibleItems] = React.useState(5);

    const handleLoadMore = () => {
        event.preventDefault()
        event.stopImmediatePropagation()
        setVisibleItems(prevVisibleItems => prevVisibleItems + 5);
    };

    return (
        <>

            <div className="row mt-5">

                {AnnouncementData && AnnouncementData.length > 0 ?
                    //.filter(x => x.FeaturedAnnouncement != false)
                    AnnouncementData.slice(0, 1).map(item => {
                        const AnnouncementandNewsBannerImage = item.AnnouncementandNewsBannerImage == undefined || item.AnnouncementandNewsBannerImage == null ? ""
                            : JSON.parse(item.AnnouncementandNewsBannerImage);
                        return (
                            <><div className="col-lg-5" >
                                <div className="imagemani mt-2 me-2">
                                    <img src={AnnouncementandNewsBannerImage?.serverUrl + AnnouncementandNewsBannerImage?.serverRelativeUrl} className="d-flex align-self-center me-3 w-100" lt="Generic placeholder image" />


                                </div>
                            </div>
                                <div className="col-lg-7">
                                    <div className="row">
                                        <div className="col-sm-4 text-left">
                                            <span style={{ padding: '5px', borderRadius: '4px', fontWeight: '500', color: '#009157', background: 'rgba(0, 135, 81, 0.20)' }} className="font-14 float-start mt-2">
                                                Latest Announcement</span>

                                        </div>
                                        <div className="col-lg-12">
                                            <h4 onClick={() => gotoAnnouncementDetails(item)} style={{ lineHeight: '34px' }} className="page-title fw-700 mb-1  pe-5 font-28 titleHeading">
                                                {item.Title}</h4>
                                        </div>
                                        <div className="row">
                                            <div className="col-sm-12">
                                                <p className="mb-2 mt-1 text-dark d-block customhead">
                                                    <span style={{ fontWeight: '400' }} className="pe-2 text-nowrap color-new font-14 mb-0 d-inline-block">
                                                        <Calendar size={12} strokeWidth={1} className="pl-2" style={{ fontWeight: '400' }} />&nbsp;&nbsp;
                                                        {moment(item.Modified).format("DD-MMM-YYYY")} &nbsp;  &nbsp;  &nbsp;|
                                                    </span>
                                                    <span style={{ fontWeight: '400' }} className="text-nowrap mb-0 color-new font-14 d-inline-block">
                                                        Author: <span style={{ color: '#009157', fontWeight: '600' }}>{item.Author.Title} &nbsp;  &nbsp;  &nbsp;
                                                        </span>

                                                    </span></p>

                                                <div style={{ clear: 'both' }}>
                                                    <p style={{ lineHeight: '20px', fontSize: '15px' }} className="d-block  text-dark customdescription">{truncateText(item.Overview, 250)}</p>
                                                </div>
                                                <a onClick={() => gotoAnnouncementDetails(item)} style={{ textDecoration: 'none' }}>
                                                    <div style={{ height: '40px', lineHeight: '24px' }} className="btnCustomcss btn-primary">Read more..</div> </a>

                                            </div>
                                        </div>
                                    </div>

                                </div></>)
                    }) :
                    <div className='row mt-2'>
                        <div style={{ height: '300px' }} className="card card-body align-items-center  annusvg text-center"
                        >

                            <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"></path><path d="M13.73 21a2 2 0 0 1-3.46 0"></path></svg>

                            <p className="font-14 text-muted text-center">No Announcement Available </p>

                        </div>
                    </div>
                }
            </div>
            <div className="tab-content mt-4">

                <div className="tab-pane show active" id="home1" role="tabpanel">
                    {AnnouncementData && AnnouncementData.length > 0 ?
                        AnnouncementData.slice(1, visibleItems).map(item => {
                            const AnnouncementandNewsBannerImage = item.AnnouncementandNewsBannerImage == undefined || item.AnnouncementandNewsBannerImage == null ? "" : JSON.parse(item.AnnouncementandNewsBannerImage);

                            return (
                                <div className="card mb-2" style={{ cursor: 'pointer' }}>
                                    <div className="card-body">
                                        <div className="row align-items-start">
                                            <div className="col-sm-2">
                                                <a onClick={() => gotoAnnouncementDetails(item)}>
                                                    <div className="imagehright">
                                                        <img src={AnnouncementandNewsBannerImage?.serverUrl + AnnouncementandNewsBannerImage?.serverRelativeUrl}
                                                            className="d-flex align-self-center me-3 w-100" alt="Generic placeholder image" style={{ objectFit: 'cover' }} />
                                                    </div>
                                                </a>
                                            </div>
                                            <div className="col-sm-9">
                                                <div className="row">
                                                    <div className="col-sm-4 date-color">
                                                        <span className="font-12 date-color float-start mt-0 mb-1 ng-binding" style={{ color: '#6b6b6b', fontSize: '12px', paddingRight: '0.2rem' }}>
                                                            <Calendar size={12} color="#6b6b6b" strokeWidth={2} style={{ fontWeight: '400' }} /></span>

                                                        <span className="font-12 date-color float-start mt-0 mb-1 ng-binding" style={{ color: '#6b6b6b', fontSize: '12px' }}>{moment(item.Modified).format("DD-MMM-YYYY")}</span>
                                                    </div>
                                                </div>
                                                <a onClick={() => gotoAnnouncementDetails(item)}>
                                                    <div className="w-100">
                                                        <h4 className="mt-1 mb-1 hovertext font-16 fw-bold ng-binding" style={{ color: '#343a40', fontSize: '16px' }}>{truncateText(item.Title, 90)}</h4>
                                                        <p style={{ color: '#6b6b6b', fontSize: '15px',height: '4rem' }} className="mb-2 ng-binding heauu">{truncateText(item.Overview, 350)}</p>
                                                        <div className="readmore mt-3 mb-0">Read more..</div>
                                                    </div>
                                                </a>
                                            </div>
                                            <div className="col-sm-1 posx">
                                                <div className="d-flex" style={{ justifyContent: 'end', cursor: 'pointer' }}>
                                                    <div className="" style={{ position: 'relative', right: '10px' }}>
                                                        <div className="" onClick={() => toggleDropdown(item.Id)} key={item.Id}>
                                                            <Share2 size={25} color="#6c757d" strokeWidth={2} style={{ fontWeight: '400' }} />
                                                        </div>
                                                        {showDropdownId === item.Id && (
                                                            <div className="dropdown-menu dropcss" isMenuOpenshareref={menuRef}>
                                                                <a className="dropdown-item dropcssItem" onClick={() => sendanEmail(item)}> <Share size={16} />  Share by email</a>
                                                                <a className="dropdown-item dropcssItem" onClick={() => copyToClipboard(item.Id)}><Link size={14} />  Copy Link</a>
                                                                <a>{copySuccess && <span className="text-success">{copySuccess}</span>}</a>
                                                            </div>
                                                        )}
                                                    </div>
                                                    {/* <Bookmark size={20} color="#6c757d" strokeWidth={2} style={{ fontWeight: '400' }} /> */}
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            )
                        }) : null}

                    {visibleItems < AnnouncementData.length && (
                        <div className="text-center">
                            <button onClick={() => handleLoadMore()} className="btn btn-primary mt-3">Load More</button>
                        </div>
                    )}
                </div>

                <div style={{ marginBottom: "50px" }}></div>
            </div>

        </>

    )
}

export default CustomWebpartTemplate