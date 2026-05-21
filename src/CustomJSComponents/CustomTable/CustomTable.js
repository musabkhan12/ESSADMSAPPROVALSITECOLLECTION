import React, { useEffect, useState } from 'react'
import "../CustomTable/CustomTable.scss";
import "bootstrap/dist/css/bootstrap.min.css";
import "../../CustomCss/mainCustom.scss";
import "../../Assets/Figtree/Figtree-VariableFont_wght.ttf";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faEdit } from '@fortawesome/free-regular-svg-icons';

const ITEMS_PER_PAGE = 5;

const CustomTable = ({ data, headers }) => {
    debugger
    const [currentPage, setCurrentPage] = useState(1);
    const itemsPerPage = 5;
    const totalPages = Math.ceil(data.length / itemsPerPage);
    // const handlePageChange = (pageNumber) => {
    //     setCurrentPage(pageNumber);
    // };
    const handlePageChange = (pageNumber) => {
        if (pageNumber > 0 && pageNumber <= totalPages) {
            setCurrentPage(pageNumber);
        }
    };

    const startIndex = (currentPage - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    const currentData = data.slice(startIndex, endIndex);

    return (
        <div className="tab-content mt-3">
            <div className="tab-pane show active" id="profile1" role="tabpanel">
                <div className="card cardCss">
                    <div className="card-body">
                        <div id="cardCollpase4" className="collapse show">
                            <div className="table-responsive pt-0">
                                <table className="mtable table-centered table-nowrap table-borderless mb-0">
                                    <thead>
                                        <tr>
                                            {headers.map((header, index) => (
                                                <th
                                                    key={index}
                                                    style={header.style}
                                                    className={header.className}
                                                >
                                                    {header.label}
                                                </th>
                                            ))}
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {currentData.map((item) => {
                                            return (
                                                <tr key={item.id}>
                                                    {headers.map((header, headerIndex) => {
                                                        const ImageUrl = header.type === 'image' ? (item[header.key] == undefined || item[header.key] == null ? "" : JSON.parse(item[header.key])) : "";
                                                        return (
                                                            <><td key={headerIndex} style={header.cellStyle} className="table-cell">
                                                                {header.type === 'image' ?

                                                                    (
                                                                        <img
                                                                            src={ImageUrl.serverUrl + ImageUrl.serverRelativeUrl}
                                                                            alt="Table Image"
                                                                            className="table-image" style={{
                                                                                objectFit: 'cover',
                                                                                width: '100%'
                                                                            }} />
                                                                    ) :
                                                                    item[header.key]}
                                                            </td>
                                                                <td>
                                                                    <a href="edit-announcement.html" className="action-icon text-primary">
                                                                        <FontAwesomeIcon icon={faEdit} />
                                                                    </a>
                                                                    <a href="javascript:void(0);" className="action-icon text-danger">

                                                                        <img src={require('../../CustomAsset/trashed.svg')} />
                                                                    </a>
                                                                </td></>

                                                        )
                                                    })}

                                                </tr>

                                            )

                                        })}
                                    </tbody>
                                </table>

                                <nav className="pagination-container">
                                    <ul className="pagination">
                                        <li className={`page-item ${currentPage === 1 ? 'disabled' : ''}`}>
                                            <a
                                                className="page-link"
                                                onClick={() => handlePageChange(currentPage - 1)}
                                                aria-label="Previous"
                                            >
                                                «
                                            </a>
                                        </li>
                                        {Array.from({ length: totalPages }, (_, num) => (
                                            <li
                                                key={num}
                                                className={`page-item ${currentPage === num + 1 ? 'active' : ''}`}
                                            >
                                                <a
                                                    className="page-link"
                                                    onClick={() => handlePageChange(num + 1)}
                                                >
                                                    {num + 1}
                                                </a>
                                            </li>
                                        ))}
                                        <li className={`page-item ${currentPage === totalPages ? 'disabled' : ''}`}>
                                            <a
                                                className="page-link"
                                                onClick={() => handlePageChange(currentPage + 1)}
                                                aria-label="Next"
                                            >
                                                »
                                            </a>
                                        </li>
                                    </ul>
                                </nav>
                            </div> {/* .table-responsive */}
                        </div> {/* end collapse */}
                    </div> {/* end card-body */}
                </div>
            </div >
        </div >
    )
}

export default CustomTable