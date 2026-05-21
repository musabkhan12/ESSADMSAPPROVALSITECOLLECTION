import React, { useEffect, useState, useRef } from 'react'
import g1 from "../../CustomAsset/Officials.jpg";
import "bootstrap/dist/css/bootstrap.min.css";
import "./CustomNewsWebpartTemplate.scss";
import "../../CustomCss/mainCustom.scss";
import "../../Assets/Figtree/Figtree-VariableFont_wght.ttf";
import { Share2 } from 'feather-icons-react';
import { Bookmark } from 'feather-icons-react';
import { Calendar, Link, Share } from 'feather-icons-react';
import moment from 'moment';
import { getNews } from '../../APISearvice/NewsService';
import { getCurrentUserProfileEmail } from "../../APISearvice/CustomService";
const CustomNewsWebpartTemplate = ({ _sp, SiteUrl }) => {

    const [copySuccess, setCopySuccess] = useState('');
    const [show, setShow] = useState(false)
    const [itemsToShow, setItemsToShow] = useState(5); // Initial number of items to show
    const [NewsData, setNews] = useState([])
    const [showDropdownId, setShowDropdownId] = useState(null);
    const [currentEmail, setEmail] = useState('');
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
        setEmail(getCurrentUserProfileEmail(_sp))
        setNews(await getNews(_sp))

    }
    const truncateText = (text, maxLength) => {
        if (text != null) {
            return text.length > maxLength ? text.substring(0, maxLength) + "..." : text;

        }
    };


    const gotoNewsDetails = (valurArr) => {
        localStorage.setItem("NewsId", valurArr.Id)
        localStorage.setItem("NewsArr", JSON.stringify(valurArr))
        setTimeout(() => {
            window.location.href = `${SiteUrl}/SitePages/NewsDetails.aspx?${valurArr.Id}`;
        }, 1000);
    }

    const copyToClipboard = (Id) => {
        const link = `${SiteUrl}/SitePages/NewsDetails.aspx?${Id}`;
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

        // const subject = "News Title-" + item.Title;
        // const body = 'Here is the link to the event:' + `${SiteUrl}/SitePages/NewsDetails.aspx?${item.Id}`;

        //const mailtoLink = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

        // Open the link to launch the default mail client (like Outlook)
        // window.location.href = mailtoLink;
        //const office365MailLink = `https://outlook.office.com/mail/deeplink/compose?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
        const subject = "Thought You’d Find This Interesting!";
        const body = 'Hi,' +
            'I came across something that might interest you: ' +
            `<a href="${siteUrl}/SitePages/NewsDetails.aspx?${item.Id}"></a>`
        'Let me know what you think!';

        // window.open(`https://outlook.office365.com/mail/deeplink/compose?body= Here is the group link: "${siteUrl}/SitePages/GroupandTeamDetails.aspx?${idNum}"`);
        const office365MailLink = `https://outlook.office.com/mail/deeplink/compose?subject=${subject}&body=${body}`;

        window.open(office365MailLink, '_blank');
    };
    const loadMore = () => {
        event.preventDefault()
        event.stopImmediatePropagation()
        setItemsToShow(itemsToShow + 5); // Increase the number by 8
    };
    return (
        <>


            <div className="row mt-5" >
                {NewsData.length > 0 ?
                    //.filter(x => x.FeaturedAnnouncement != false)
                    NewsData.slice(0, 1).map(item => {
                        const AnnouncementandNewsBannerImage = item.AnnouncementandNewsBannerImage == undefined || item.AnnouncementandNewsBannerImage == null ? ""
                            : JSON.parse(item.AnnouncementandNewsBannerImage);
                        return (
                            <><div className="col-lg-5">
                                <div className="imagemani mt-2 me-2">
                                    <img src={AnnouncementandNewsBannerImage?.serverUrl + AnnouncementandNewsBannerImage?.serverRelativeUrl}
                                        className="d-flex align-self-center me-3 w-100" lt="Generic placeholder image" style={{ objectFit: 'cover' }} />


                                </div>
                            </div>
                                <div className="col-lg-7">
                                    <div className="row" style={{ paddingLeft: '0.5rem' }}>
                                        <div className="col-sm-4 text-left">
                                            <span style={{ padding: '5px', borderRadius: '4px', fontWeight: '500', color: '#009157', background: 'rgba(0, 135, 81, 0.20)' }} className="font-14 float-start mt-2">
                                                Latest News</span>

                                        </div>
                                        <div className="col-lg-12">
                                            <h4 onClick={() => gotoNewsDetails(item)} style={{ lineHeight: '34px' }} className="page-title fw-700 mb-1  pe-5 font-28 titleHeading">
                                                {item.Title}</h4>
                                        </div>
                                        <div className="row">
                                            <div className="col-sm-12">
                                                <p className="mb-2 mt-1 d-block text-dark customhead">
                                                    <span style={{ fontWeight: '400' }} className="pe-2 text-nowrap font-14 mb-0 d-inline-block">
                                                        <Calendar size={12} strokeWidth={1} className="pl-2 text-muted" style={{ fontWeight: '400' }} />&nbsp;
                                                        {moment(item.Modified).format("DD-MMM-YYYY")} &nbsp;  &nbsp;  &nbsp;|
                                                    </span>
                                                    <span style={{ fontWeight: '400' }} className="text-nowrap mb-0  font-14 text-muted2 d-inline-block">
                                                        Author: <span style={{ color: '#009157', fontWeight: '600' }}>{item.Author.Title}
                                                        </span>

                                                    </span></p>

                                                <div style={{ clear: 'both' }}>
                                                    <p style={{ lineHeight: '22px', fontSize: '15px' }} className="d-block  text-dark customdescription">{truncateText(item.Overview, 320)}</p>
                                                </div>
                                                <a onClick={() => gotoNewsDetails(item)} style={{ textDecoration: 'none', marginTop: '15px', float: 'left' }}>
                                                    <div style={{ height: '40px', lineHeight: '24px' }} className="btnCustomcss btn-primary">Read more..</div> </a>

                                            </div>
                                        </div>
                                    </div>

                                </div></>)
                    }) :
                    <div className='row mt-2'>
                        <div style={{ height: '300px' }} className="card card-body align-items-center  annusvg text-center"
                        >

                            <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"></path><polyline points="13 2 13 9 20 9"></polyline></svg>

                            <p className="font-14 text-muted text-center">No News Available </p>

                        </div>
                    </div>
                }
            </div>
            <div className="tab-content mt-4">

                <div className="tab-pane show active" id="home1" role="tabpanel">
                    {NewsData.length > 0 ?
                        NewsData.slice(1, itemsToShow).map(item => {
                            const AnnouncementandNewsBannerImage = item.AnnouncementandNewsBannerImage == undefined || item.AnnouncementandNewsBannerImage == null ? "" : JSON.parse(item.AnnouncementandNewsBannerImage);


                            return (
                                <div className="card mb-2 annuncementcard" style={{ cursor: 'pointer' }}>
                                    <div className="card-body">
                                        <div className="row align-items-start">
                                            <div className="col-sm-2">
                                                <a onClick={() => gotoNewsDetails(item)}>   <div className="imagehright">
                                                    {/* <img className="d-flex align-self-center me-3 w-100" src={g1} alt="Generic placeholder image" /> */}
                                                    <img src={AnnouncementandNewsBannerImage?.serverUrl + AnnouncementandNewsBannerImage?.serverRelativeUrl}
                                                        className="d-flex align-self-center me-3 w-100" lt="Generic placeholder image" style={{ height: '100%' }} />
                                                </div>
                                                </a>
                                            </div>
                                            <div className="col-sm-9">
                                                <div className="row">
                                                    <div className="col-sm-4 date-color">
                                                        <span className="font-12 date-color float-start mt-0 mb-1 ng-binding" style={{ color: '#6b6b6b', fontSize: '12px', paddingRight: '0.2rem' }}>
                                                            <Calendar size={12} color="#6b6b6b" strokeWidth={2} style={{ fontWeight: '400' }} /></span>

                                                        <span className="font-12 date-color float-start mt-0 mb-1 ng-binding" style={{ color: '#6b6b6b', fontSize: '12px' }}>{moment(item.Modified).format("DD-MMM-YYYY")}
                                                            {/* 12-Mar-2024 18:37 */}
                                                        </span>
                                                    </div>
                                                </div>
                                                <div className="w-100">
                                                    <a >  <h4 onClick={() => gotoNewsDetails(item)} className="hovertext mt-1 mb-1 font-16 fw-bold ng-binding" style={{ color: '#343a40', fontSize: '16px' }}> {truncateText(item.Title, 80)}
                                                    </h4> </a>
                                                    <p style={{ color: '#6b6b6b', fontSize: '15px', cursor: 'auto', height: '4rem' }} className="mb-2  ng-binding parah heauu">
                                                        {truncateText(item.Overview, 250)}</p>
                                                    <div className="readmore" onClick={() => gotoNewsDetails(item)} style={{ cursor: 'pointer' }}>Read more..</div>
                                                </div>

                                            </div>
                                            <div className="col-sm-1 posx">
                                                <div className="d-flex" style={{ justifyContent: 'end', marginRight: "10px", cursor: 'pointer' }}>
                                                    <div className="" style={{ position: 'relative' }}>
                                                        <div className="" onClick={() => toggleDropdown(item.Id)} key={item.Id}>
                                                            <Share2 size={25} color="#6c757d" strokeWidth={2} style={{ fontWeight: '400' }} />
                                                        </div>
                                                        {showDropdownId === item.Id && (
                                                            <div className="dropdown-menu dropcss" isMenuOpenshareref={menuRef}>
                                                                <a className="dropdown-item dropcssItem" onClick={() => sendanEmail(item)}><Share size={16} />  Share by email</a>
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
                        }
                        ) : null}

                </div>
                {itemsToShow < NewsData.length && (
                    <div className="col-12 text-center mt-3">
                        <button onClick={loadMore} className="btn btn-primary">
                            Load More
                        </button>
                    </div>
                )}
            </div>
            <div style={{height:'70px'}}></div>
            
            </>

    )
}

export default CustomNewsWebpartTemplate