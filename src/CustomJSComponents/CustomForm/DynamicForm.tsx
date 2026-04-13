import * as React from 'react';
import "bootstrap/dist/css/bootstrap.min.css";
import 'react-quill/dist/quill.snow.css';
import ReactQuill from 'react-quill';
import { uploadFileToLibrary } from '../../APISearvice/AnnouncementsService';
import "../../CustomCss/mainCustom.scss";
import "../CustomForm/CustomForm.scss";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faDeleteLeft, faPaperclip, faRemove } from '@fortawesome/free-solid-svg-icons';
import { Modal } from 'react-bootstrap';
import SweetAlert from 'react-bootstrap-sweetalert';
import  { fireAlert } from "../SweetAlert2Components/SweetAlert2";
import Swal from 'sweetalert2';

interface FormField {
    type: string;
    name: string;
    label: string;
    docLib?: string;
    placeholder?: string;
    options?: any[];
    required: boolean;
}

interface DynamicFormProps {
    formSchemas: FormField[];
    onChange: (name: string, value: string) => void;
    onSubmit: (formValues: { [key: string]: string }, BnnerImagepostArr: any[], ImagepostArr: any[], DocumentpostArr: any[]) => void;
    sp: any;
    Spurl: any;
    onCancel: (boleanVal: boolean) => void;
}

const DynamicForm: React.FC<DynamicFormProps> = ({ formSchemas, onChange, onSubmit, sp, Spurl, onCancel }) => {

    const [DocsArr, setDocsArr] = React.useState([]);
    const [GalaryArr, setGalaryArr] = React.useState([]);
    const [richTextValues, setRichTextValues] = React.useState<{ [key: string]: string }>({});
    const [DocumentpostArr, setDocumentpostArr] = React.useState([]);
    const [ImagepostArr, setImagepostArr] = React.useState([]);
    const [BnnerImagepostArr, setBannerImagepostArr] = React.useState([]);
    const [showModal, setShowModal] = React.useState(false);
    const [showDocTable, setShowDocTable] = React.useState(false);
    const [showImgModal, setShowBannerTable] = React.useState(false);
    const [showBannerModal, setShowImgTable] = React.useState(false);
    const [showAlert, setShowAlert] = React.useState(false);
    const [alertTitle, setAlertTitle] = React.useState('');
    const [alertType, setAlertType] = React.useState<'success' | 'error' | 'warning'>('success');
  



    const handleSubmit = (e: React.FormEvent) => {
        e.preventDefault();
        const formValues: { [key: string]: string } = {};
        // Show success alert after submission
        Swal.fire({
            title: 'Do you want to save',
            showConfirmButton: true,
            showCancelButton: true,
            confirmButtonText: "Save",
            cancelButtonText: "Cancel",
            icon: 'warning'
        }
        ).then((result) => {
            /* Read more about isConfirmed, isDenied below */
            if (result.isConfirmed) {
                formSchemas.forEach(field => {
                    const input = (e.target as HTMLFormElement).elements.namedItem(field.name);
        
                    if (input && input instanceof HTMLInputElement && input.type === 'file') {
                        // Handle file input (get file names)
                        const files = input.files;
                        formValues[field.name] = files ? Array.apply(files).map((file: { name: any; }) => file.name).join(', ') : "";
                    } else if (input instanceof HTMLInputElement || input instanceof HTMLTextAreaElement || input instanceof HTMLSelectElement) {
                        formValues[field.name] = input.value;
                    } else if (field.type === 'richtext') {
                        formValues[field.name] = richTextValues[field.name] || ""; // Get value from state
                    }
                });
                onSubmit(formValues, BnnerImagepostArr, ImagepostArr, DocumentpostArr);
                Swal.fire('Nice to meet you', '', 'success');

            } else
            {
                Swal.fire(' Cancelled', '', 'error')
            }
              

        })
     
       

    //  // Show success alert after submission
    //  setAlertTitle('Form Submitted Successfully!');
    //  setAlertType('success');
    //  setShow//alert(true);
   };
 
//    const hideAlert = () => {
//      setShow.alert(false);
//    };

    const handleCancel = () => {
        onCancel(true)
    }

    const onFileChange = async (event: React.ChangeEvent<HTMLInputElement>, libraryName: string, docLib: string) => {
        debugger;
        event.preventDefault();
        let uloadDocsFiles: any[] = [];
        let uloadImageFiles: any[] = [];
        let uloadBannerImageFiles: any[] = [];

        if (event.target.files && event.target.files.length > 0) {
            const files = Array.from(event.target.files);

            if (libraryName === "docs") {
                const docFiles = files.filter(file =>
                    file.type === 'application/pdf' ||
                    file.type === 'application/msword' ||
                    file.type === 'application/xsls' || file.type === 'text/csv'
                );

                if (docFiles.length > 0) {
                    const arr = {
                        files: docFiles,
                        libraryName: libraryName,
                        docLib: docLib
                    };
                    uloadDocsFiles.push(arr);
                    setDocumentpostArr(uloadDocsFiles);
                } else {
                    //alert("only docuemnt can be upload")
                }
            }
            if (libraryName === "Gallery" || libraryName === "bannerimg") {
                const imageVideoFiles = files.filter(file =>
                    file.type.startsWith('image/') ||
                    file.type.startsWith('video/')
                );

                if (imageVideoFiles.length > 0) {
                    const arr = {
                        files: imageVideoFiles,
                        libraryName: libraryName,
                        docLib: docLib
                    };
                    if (libraryName === "Gallery") {
                        uloadImageFiles.push(arr);
                        setImagepostArr(uloadImageFiles);
                    } else {
                        uloadBannerImageFiles.push(arr);
                        setBannerImagepostArr(uloadBannerImageFiles);
                    }
                } else {
                    //alert("only image & video can be upload")
                }
            }
        }
    };

    const deleteLocalFile = (index: number, filArray: any[], title: string, name: string) => {
        debugger
        // Remove the file at the specified index
        filArray[0].files.splice(index, 1);

        // Update the state based on the title
        if (name === "bannerimg") {
            setBannerImagepostArr([...filArray]);
            filArray[0].files.length > 0 ? "" : setShowModal(false); clearFileInput(name);
        } else if (name === "Gallery") {
            setImagepostArr([...filArray]);
            filArray[0].files.length > 0 ? "" : setShowModal(false); clearFileInput(name);
        } else {
            setDocumentpostArr([...filArray]);
            filArray[0].files.length > 0 ? "" : setShowModal(false); clearFileInput(name);
        }
        // Clear the file input

    };
    const clearFileInput = (name: any) => {
        debugger
        const input = document.querySelector(`input[name=${name}]`) as HTMLInputElement;
        if (input) {
            input.value = ''; // Clears the selected files
        }
        else if (input == null) {
            input.value = ''; // Clears the selected files
        }
    };

    const onUpload = async (files: FileList, libraryname: string, docLib: string) => {
        let uploadedFilesArr = [];
        if (libraryname === "AnnouncementAndNewsDocs") {
            if (files.length > 0) {
                // const arr = await uploadFileToLibrary(files, libraryname, sp, docLib);
                //   setDocsArr(prevState => [...prevState, ...arr]);
            }
        } else {
            if (files.length > 0) {
                // const arr = await uploadFileToLibrary(files, libraryname, sp, docLib);
                //setGalaryArr(prevState => [...prevState, ...arr]);
            }
        }
    };

    const formats = [
        "header", "height", "bold", "italic",
        "underline", "strike", "blockquote",
        "list", "color", "bullet", "indent",
        "link", "image", "align", "size",
    ];

    const modules = {
        toolbar: [
            [{ size: ["small", false, "large", "huge"] }],
            ["bold", "italic", "underline", "strike", "blockquote"],
            [{ list: "ordered" }, { list: "bullet" }],
            ["link", "image"],
            [
                { list: "ordered" },
                { list: "bullet" },
                { indent: "-1" },
                { indent: "+1" }
            ],
            [{ "color": ["#000000", "#e60000", "#ff9900", "#ffff00", "#008a00", "#0066cc", "#9933ff", "#ffffff", "#facccc", "#ffebcc", "#ffffcc", "#cce8cc", "#cce0f5", "#ebd6ff", "#bbbbbb", "#f06666", "#ffc266", "#ffff66", "#66b966", "#66a3e0", "#c285ff", "#888888", "#a10000", "#b26b00", "#b2b200", "#006100", "#0047b2", "#6b24b2", "#444444", "#5c0000", "#663d00", "#666600", "#003700", "#002966", "#3d1466", 'custom-color'] }],
        ]
    };

    console.log(DocumentpostArr, 'DocumentpostArr', ImagepostArr, 'ImagepostArr', BnnerImagepostArr, 'BnnerImagepostArr');
    const setShowModalFunc = (bol: boolean, name: string) => {
        debugger
        if (name == "bannerimg") {
            setShowModal(bol)
            setShowBannerTable(true)
            setShowImgTable(false)
            setShowDocTable(false)
        }
        else if (name == "Gallery") {
            setShowModal(bol)
            setShowImgTable(true)
            setShowBannerTable(false)
            setShowDocTable(false)
        }
        else {
            setShowModal(bol)
            setShowDocTable(true)
            setShowBannerTable(false)
            setShowImgTable(false)
        }
    }

    return (
        <div className="card cardcss mt-3">
            <div className="card-body">
                <div className="row mt-2">

                    <form className='row' onSubmit={handleSubmit} onTouchCancel={handleCancel}>
                        {formSchemas.map((field, index) => (
                            <>
                                {field.type === 'select' ? (
                                    <div key={index} className="col-md-4 mb-3">
                                        <label className="form-label">
                                            {field.label} {field.required && <span className="text-danger">*</span>}
                                        </label>
                                        <select
                                            className="form-select inputcss"
                                            name={field.name}
                                            onChange={(e) => onChange(field.name, e.target.value)}
                                            required={field.required}
                                        >
                                            <option value="">Select</option>
                                            {field.options?.map(option => (
                                                <option key={option.id} value={option.id}>
                                                    {option.name}
                                                </option>
                                            ))}
                                        </select>
                                    </div>
                                ) : field.type === 'textarea' ? (
                                    <div key={index} className="col-md-4 mb-3">
                                        <label className="form-label">
                                            {field.label} {field.required && <span className="text-danger">*</span>}
                                        </label>
                                        <textarea
                                            className="form-control inputcss"
                                            name={field.name}
                                            placeholder={field.placeholder}
                                            rows={5}
                                            onChange={(e) => onChange(field.name, e.target.value)}
                                            required={field.required}
                                            style={{ height: '100px' }}
                                        />
                                    </div>
                                ) : field.type === 'file' ? (
                                    <div key={index} className="col-md-4 mb-3">
                                        <div className='d-flex justify-content-between'>
                                            <div>
                                                <label className="form-label" >
                                                    {field.label}
                                                </label>
                                            </div>
                                            <div>
                                                {field.docLib == "ARGAnnouncementAndNewsDocs" ? DocumentpostArr.length > 0 &&
                                                    DocumentpostArr[0].files.length == 1 &&
                                                    (<a onClick={() => setShowModalFunc(true, "docs")} style={{ fontSize: '0.875rem' }}>
                                                        <FontAwesomeIcon icon={faPaperclip} /> {DocumentpostArr[0].files.length} file Attached
                                                    </a>)
                                                    || DocumentpostArr.length > 0 && DocumentpostArr[0].files.length > 1 &&
                                                    (<a onClick={() => setShowModalFunc(true, "docs")} style={{ fontSize: '0.875rem' }}>
                                                        <FontAwesomeIcon icon={faPaperclip} /> {DocumentpostArr[0].files.length} files Attached
                                                    </a>)
                                                    : null}
                                                {field.docLib == "AnnouncementAndNewsGallary" ? ImagepostArr.length > 0 &&
                                                    ImagepostArr[0].files.length == 1 &&
                                                    (<a onClick={() => setShowModalFunc(true, field.name)} style={{ fontSize: '0.875rem' }}>
                                                        <FontAwesomeIcon icon={faPaperclip} /> {ImagepostArr[0].files.length} file Attached
                                                    </a>)
                                                    || ImagepostArr.length > 0 && ImagepostArr[0].files.length > 1 &&
                                                    (<a onClick={() => setShowModalFunc(true, field.name)} style={{ fontSize: '0.875rem' }}>
                                                        <FontAwesomeIcon icon={faPaperclip} /> {ImagepostArr[0].files.length} files Attached
                                                    </a>)
                                                    : null}
                                                {field.docLib == "AnnouncementandNewsBannerImage" ? BnnerImagepostArr.length > 0 &&
                                                    BnnerImagepostArr[0].files.length == 1 &&
                                                    (<a onClick={() => setShowModalFunc(true, field.name)} style={{ fontSize: '0.875rem' }}>
                                                        <FontAwesomeIcon icon={faPaperclip} /> {BnnerImagepostArr[0].files.length} file Attached
                                                    </a>)
                                                    || BnnerImagepostArr.length > 0 && BnnerImagepostArr[0].files.length > 1 &&
                                                    (<a onClick={() => setShowModalFunc(true, field.name)} style={{ fontSize: '0.875rem' }}>
                                                        <FontAwesomeIcon icon={faPaperclip} /> {BnnerImagepostArr[0].files.length} files Attached
                                                    </a>)
                                                    : null}
                                            </div>
                                        </div>
                                        {field.docLib == "ARGAnnouncementAndNewsDocs" && <input
                                            className="form-control inputcss"
                                            type="file"
                                            name={field.name}
                                            multiple
                                            onChange={(e) => onFileChange(e, field.name, field.docLib)}
                                            required={field.required}
                                        />}
                                        {field.docLib == "AnnouncementAndNewsGallary" && <input
                                            className="form-control inputcss"
                                            type="file"
                                            name={field.name}
                                            multiple
                                            onChange={(e) => onFileChange(e, field.name, field.docLib)}
                                            required={field.required}
                                        />}
                                        {field.docLib == "AnnouncementandNewsBannerImage" && <input
                                            className="form-control inputcss"
                                            type="file"
                                            name={field.name}
                                            onChange={(e) => onFileChange(e, field.name, field.docLib)}
                                            required={field.required}
                                        />}
                                    </div>
                                ) : field.type === 'radio' ? (
                                    <div key={index} className="col-md-3 mb-3">
                                        <label className="form-label">
                                            {field.label} {field.required && <span className="text-danger">*</span>}
                                        </label>
                                        {field.options?.map(option => (
                                            <div key={option.id} className="form-check">
                                                <input
                                                    className="form-check-input inputcss"
                                                    type="radio"
                                                    name={field.name}
                                                    value={option.id}
                                                    onChange={(e) => onChange(field.name, e.target.value)}
                                                    required={field.required}
                                                />
                                                <label className="form-check-label">{option.name}</label>
                                            </div>
                                        ))}
                                    </div>
                                ) : field.type === 'richtext' ? (
                                    <div key={index} className="col-md-5 mb-3">
                                        <label className="form-label">
                                            {field.label} {field.required && <span className="text-danger">*</span>}
                                        </label>

                                        <div style={{ display: "grid", justifyContent: "start" }}>
                                            <ReactQuill
                                                theme="snow"
                                                modules={modules}
                                                formats={formats}
                                                placeholder={field.placeholder || 'Write your content ...'}
                                                onChange={(content) => {
                                                    onChange(field.name, content);
                                                    setRichTextValues({ ...richTextValues, [field.name]: content });
                                                }}
                                                style={{ width: '100%', fontSize: '6px' }}
                                            />
                                        </div>
                                    </div>
                                ) : field.type === 'checkbox' ? (
                                    <div key={index} className="col-md-3 mb-3">
                                        <label className="form-label">
                                            {field.label} {field.required && <span className="text-danger">*</span>}
                                        </label>
                                        <div className="form-check">
                                            <input
                                                className="form-check-input inputcss"
                                                type="checkbox"
                                                name={field.name}
                                                onChange={(e) => onChange(field.name, e.target.checked ? 'true' : 'false')}
                                                required={field.required}
                                            />
                                        </div>
                                    </div>
                                ) : (
                                    <div key={index} className="col-md-4 mb-3">
                                        <label className="form-label">
                                            {field.label} {field.required && <span className="text-danger">*</span>}
                                        </label>
                                        <input
                                            className="form-control inputcss"
                                            type={field.type}
                                            name={field.name}
                                            placeholder={field.placeholder}
                                            onChange={(e) => onChange(field.name, e.target.value)}
                                            required={field.required}
                                        />
                                    </div>
                                )}
                            </>
                        ))}
                        <div className="text-center butncss">
                            <button type="submit" className="btn btn-success waves-effect waves-light m-1" style={{ fontSize: '0.875rem' }}>
                                <div className='d-flex' style={{ justifyContent: 'space-around', width: '70px' }}>
                                    <img src={require('../../Assets/ExtraImage/checkcircle.svg')} style={{ width: '1rem' }} alt="Check" /> Submit
                                </div>
                            </button>
                            <button type="button" className="btn btn-light waves-effect waves-light m-1" style={{ fontSize: '0.875rem' }} onClick={handleCancel}>
                                <img src={require('../../Assets/ExtraImage/xIcon.svg')} style={{ width: '1rem' }}
                                    className='me-1' alt="x" />
                                Cancel
                            </button>
                        </div>
                    </form>
                </div>
            </div>

            {/* Modal to display uploaded files */}
            <Modal show={showModal} onHide={() => setShowModal(false)}>
                <Modal.Header closeButton>
                    {DocumentpostArr.length > 0 && showDocTable && <Modal.Title>Documents</Modal.Title>}
                    {ImagepostArr.length > 0 && showImgModal && <Modal.Title>Gallery Images/Videos</Modal.Title>}
                    {BnnerImagepostArr.length > 0 && showBannerModal && <Modal.Title>Banner Images</Modal.Title>}
                </Modal.Header>
                <Modal.Body>

                    {DocumentpostArr.length > 0 && showDocTable &&
                        (
                            <>
                                <table className="table table-bordered" style={{ fontSize: '0.75rem' }}>
                                    <thead style={{ background: '#eef6f7' }}>
                                        <tr>
                                            <th>Serial No.</th>
                                            <th>File Name</th>
                                            <th>File Size</th>
                                            <th className='text-center'>Action</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {DocumentpostArr[0].files.map((file: any, index: number) => (
                                            <tr key={index}>
                                                <td className='text-center'>{index + 1}</td>
                                                <td>{file.name}</td>
                                                <td className='text-right'>{file.size}</td>
                                                <td className='text-center'> <img src={require("../../CustomAsset/trashed.svg")} style={{ width: '15px' }} onClick={() => deleteLocalFile(index, DocumentpostArr, DocumentpostArr[0].docLib, "docs")} /> </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table></>
                        )
                    }
                    {ImagepostArr.length > 0 && showImgModal &&
                        (
                            <>
                                <table className="table table-bordered" style={{ fontSize: '0.75rem' }}>
                                    <thead style={{ background: '#eef6f7' }}>
                                        <tr>
                                            <th>Serial No.</th>
                                            <th>File Name</th>
                                            <th>File Size</th>
                                            <th className='text-center'>Action</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {ImagepostArr[0].files.map((file: any, index: number) => (
                                            <tr key={index}>
                                                <td className='text-center'>{index + 1}</td>
                                                <td>{file.name}</td>
                                                <td className='text-right'>{file.size}</td>
                                                <td className='text-center'> <img src={require("../../CustomAsset/trashed.svg")} style={{ width: '15px' }} onClick={() => deleteLocalFile(index, ImagepostArr, ImagepostArr[0].docLib, "Gallery")} /> </td>

                                            </tr>
                                        ))}
                                    </tbody>
                                </table></>
                        )}
                    {BnnerImagepostArr.length > 0 && showBannerModal &&
                        (
                            <>
                                <table className="table table-bordered" style={{ fontSize: '0.75rem' }}>
                                    <thead style={{ background: '#eef6f7' }}>
                                        <tr>
                                            <th>Serial No.</th>
                                            <th>File Name</th>
                                            <th>File Size</th>
                                            <th className='text-center'>Action</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {BnnerImagepostArr[0].files.map((file: any, index: number) => (
                                            <tr key={index}>
                                                <td className='text-center'>{index + 1}</td>
                                                <td>{file.name}</td>
                                                <td className='text-right'>{file.size}</td>
                                                <td className='text-center'> <img src={require("../../CustomAsset/trashed.svg")} style={{ width: '15px' }} onClick={() => deleteLocalFile(index, BnnerImagepostArr, BnnerImagepostArr[0].docLib, "bannerimg")} /> </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table></>
                        )}
                </Modal.Body>

            </Modal>

            {/* SweetAlert */}
        
        </div>

    );
};

export default DynamicForm;
