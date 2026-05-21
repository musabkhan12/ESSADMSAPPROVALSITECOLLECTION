import * as React from "react";
import { initializeIcons } from "@fluentui/react/lib/Icons";
import styles from "./NewRequest.module.scss";

initializeIcons();

import {
  TextField,
  Dropdown,
  DatePicker,
  IconButton,
  PrimaryButton,
  DefaultButton,
  IDropdownOption,
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import Select from "react-select";
import CustomPopup from "./CustomPopup";
import "./Newreg.css";
import { ErrorLogger } from "../../../utils/ErrorLogger";

import loaderGif from "../assets/Loder.gif";

interface INewRequestProps {
  context: WebPartContext;
  onCancel?: () => void;
  initialData?: any;
}

interface UserOption {
  value: number;
  label: string;
  email: string;
}

interface AreaMaster {
  ID: number;
  Area: string;
  AreaIdentificationCode: string;
  AreaCode: string;
}

interface OrganizationMaster {
  ID: number;
  Organization: string;
  OrganizationIdentificationCode: string;
}

interface DocumentTypeMaster {
  ID: number;
  DocumentType: string;
  DocumentTypeCode: string;
}

interface DeliverablesMaster {
  ID: number;
  Deliverables: string;
}

interface FolderCreationResult {
  projectFolderUrl: string;
  deliverablesFolders: Array<{
    deliverableName: string;
    folderUrl: string;
    deliverableId: number;
  }>;
}

const NewRequest: React.FC<INewRequestProps> = (props) => {
  const sp: SPFI = spfi().using(SPFx(props.context));
  const [projectTypeOptions, setProjectTypeOptions] = React.useState<
    IDropdownOption[]
  >([]);
  const [AreaOptions, setAreaOptions] = React.useState<IDropdownOption[]>([]);
  const [DeliverablesOptions, setDeliverablesOptions] = React.useState<
    IDropdownOption[]
  >([]);
  const [OrganizationOptions, setOrganizationOptions] = React.useState<
    IDropdownOption[]
  >([]);
  const [DocumentTypeOptions, setDocumentTypeOptions] = React.useState<
    IDropdownOption[]
  >([]);
  const [userOptions, setUserOptions] = React.useState<UserOption[]>([]);
  const [selectedUser, setSelectedUser] = React.useState<UserOption | null>(
    null,
  );

  // Master data arrays
  const [areaMasterArr, setAreaMasterArr] = React.useState<AreaMaster[]>([]);
  const [organizationMasterArr, setOrganizationMasterArr] = React.useState<
    OrganizationMaster[]
  >([]);
  const [documentsTypeMasterArr, setDocumentsTypeMasterArr] = React.useState<
    DocumentTypeMaster[]
  >([]);
  const [deliverablesMasterListArr, setDeliverablesMasterListArr] =
    React.useState<DeliverablesMaster[]>([]);

  // New state variables for popup and validation
  const [showPopup, setShowPopup] = React.useState(false);
  const [popupConfig, setPopupConfig] = React.useState({
    type: "confirmation" as "confirmation" | "validation" | "success" | "error",
    title: "",
    message: "",
  });
  const [fieldErrors, setFieldErrors] = React.useState<{
    [key: string]: boolean;
  }>({});

  // Created object the handel on change in project info field
  const [projectInfo, setProjectInfo] = React.useState({
    projectName: "",
    clientName: "",
    preparedBy: "",
    startDate: undefined as Date | undefined,
    projectType: "",
    overview: "",
  });

  const [showSubmitLoader, setShowSubmitLoader] =
    React.useState<boolean>(false);

    const [submitAction, setSubmitAction] = React.useState<"submit" | "draft">("submit");

    const [isDraftMode, setIsDraftMode] = React.useState(false);

  // Define deliverable item type
  interface DeliverableItem {
    deliverable: string;
    area: string;
    organisation: string;
    docType: string;
    seqNumber: string;
    docNumber: string;
    dueDate: Date | undefined;
    assignedTo: number;
  }

  // this is used for dynamic deleivreables table generation array
  const [deliverables, setDeliverables] = React.useState<DeliverableItem[]>([
    {
      deliverable: "",
      area: "",
      organisation: "",
      docType: "",
      seqNumber: "",
      docNumber: "",
      dueDate: undefined,
      assignedTo: 0,
    },
  ]);

 const validateForm = (actionType?: "submit" | "draft"): boolean => {
    const errors: { [key: string]: boolean } = {};
    let errorMessages: string[] = [];

    // Validate Project Information
   if (actionType === "draft") {

  const draftErrors: { [key: string]: boolean } = {};

  if (!projectInfo.projectName.trim()) {
  draftErrors.projectName = true;
}

if (!projectInfo.projectType) {
  draftErrors.projectType = true;
}

  setFieldErrors(draftErrors);

  if (Object.keys(draftErrors).length > 0) {

    setPopupConfig({
      type: "validation",
      title: "Validation Error",
      message: "Please fill mandatory fields.",
    });

    setShowPopup(true);
    setShowSubmitLoader(false);

    return false;
  }

  return true;
}
  // Submit mode - ALL fields mandatory (existing logic)
  if (!projectInfo.projectName.trim()) {
    errors.projectName = true;
    errorMessages.push("Project Name");
  }
  if (!projectInfo.clientName.trim()) {
    errors.clientName = true;
    errorMessages.push("Client Name");
  }
  if (!selectedUser) {
    errors.preparedBy = true;
    errorMessages.push("Prepared By");
  }
  if (!projectInfo.startDate) {
    errors.startDate = true;
    errorMessages.push("Project Start Date");
  }
  if (!projectInfo.projectType) {
    errors.projectType = true;
    errorMessages.push("Project Type");
  }
  if (!projectInfo.overview.trim()) {
    errors.overview = true;
    errorMessages.push("Project Overview");
  }
  // Validate Deliverables - at least one row must have deliverable selected
  const hasAtLeastOneDeliverable = deliverables.some(
    (d) => d.deliverable.trim() !== "",
  );
  if (!hasAtLeastOneDeliverable) {
    errors.minimumDeliverable = true;
    errorMessages.push("At least one Deliverable");
  }
  // Validate each deliverable row - ALL fields mandatory if any field is filled
  deliverables.forEach((deliverable, index) => {
    const isLastSingleRow =
      deliverables.length === 1 &&
      index === 0 &&
      !deliverable.deliverable &&
      !deliverable.area &&
      !deliverable.organisation &&
      !deliverable.docType &&
      !deliverable.dueDate &&
      !deliverable.assignedTo;
    const shouldValidateRow = !isLastSingleRow;
    if (shouldValidateRow) {
      if (!deliverable.deliverable) {
        errors[`deliverable_${index}`] = true;
        errorMessages.push(`Deliverable in row ${index + 1}`);
      }
      if (!deliverable.area) {
        errors[`area_${index}`] = true;
        errorMessages.push(`Area in row ${index + 1}`);
      }
      if (!deliverable.organisation) {
        errors[`organisation_${index}`] = true;
        errorMessages.push(`Organisation in row ${index + 1}`);
      }
      if (!deliverable.docType) {
        errors[`docType_${index}`] = true;
        errorMessages.push(`Document Type in row ${index + 1}`);
      }
      if (!deliverable.dueDate) {
        errors[`dueDate_${index}`] = true;
        errorMessages.push(`Due Date in row ${index + 1}`);
      }
      if (!deliverable.assignedTo) {
        errors[`assignedTo_${index}`] = true;
        errorMessages.push(`Assigned To in row ${index + 1}`);
      }
    }
  });
  setFieldErrors(errors);
  if (Object.keys(errors).length > 0) {
    setPopupConfig({
      type: "validation",
      title: "Validation Error",
      message: `Please fill mandatory fields.`,
    });
    setShowPopup(true);
    setShowSubmitLoader(false);
    return false;
  }

    return true;
  };

  // Helper function to find user option (alternative to Array.find)
  const findUserOption = (
    users: UserOption[],
    userId: number,
  ): UserOption | null => {
    for (let i = 0; i < users.length; i++) {
      if (users[i].value === userId) {
        return users[i];
      }
    }
    return null;
  };

  // Helper function to get master data item by text value
  const getAreaByText = (areaText: string): AreaMaster | undefined => {
    return areaMasterArr.find((area) => area.Area === areaText);
  };

  const getOrganizationByText = (
    orgText: string,
  ): OrganizationMaster | undefined => {
    return organizationMasterArr.find((org) => org.Organization === orgText);
  };

  const getDocumentTypeByText = (
    docTypeText: string,
  ): DocumentTypeMaster | undefined => {
    return documentsTypeMasterArr.find(
      (doc) => doc.DocumentType === docTypeText,
    );
  };

  const getDeliverableByText = (
    deliverableText: string,
  ): DeliverablesMaster | undefined => {
    return deliverablesMasterListArr.find(
      (del) => del.Deliverables === deliverableText,
    );
  };

  // Sequential number and document number generation function
  const generateSequenceAndDocNumber = async (
    index: number,
    source: "Project Type" | "Document Type",
  ) => {
    const row = deliverables[index];

    if (
      !projectInfo.projectName ||
      !row.area ||
      !row.organisation ||
      !row.docType
    ) {
      return;
    }

    const selectedArea = getAreaByText(row.area);
    const selectedOrganization = getOrganizationByText(row.organisation);
    const selectedDocType = getDocumentTypeByText(row.docType);
    const selectedDeliverable = getDeliverableByText(row.deliverable);

    if (!selectedArea || !selectedOrganization || !selectedDocType) {
      console.warn("Could not find matching master data");
      return;
    }

    try {
      const existingItems = await sp.web.lists
        .getByTitle("DeliverablesDetails")
        .items.select(
          "SequenticalNumber",
          "ProjectCreationListID/ID",
          "ProjectCreationListID/ProjectName",
        )
        .expand("ProjectCreationListID")
        .filter(
          `ProjectCreationListID/ProjectName eq '${projectInfo.projectName}' 
                and AreaId eq ${selectedArea.ID} 
                and OrganisationId eq ${selectedOrganization.ID} 
                and DocumentTypeId eq ${selectedDocType.ID}`,
        )
        .orderBy("Created", false)
        .top(4999)();

      // Get max sequence from SharePoint list
      const maxSeqFromList =
        existingItems.length > 0
          ? Math.max(
              ...existingItems.map((r) =>
                r.SequenticalNumber ? parseInt(r.SequenticalNumber) : 0,
              ),
            )
          : 0;

      // Get max sequence from current deliverables array excluding current index
      const maxSeqFromArray = Math.max(
        ...deliverables.map((r, i) => {
          if (i === index) return 0; // skip current row

          const isMatch =
            r.area === row.area &&
            r.organisation === row.organisation &&
            r.docType === row.docType;

          return isMatch && r.seqNumber ? parseInt(r.seqNumber) : 0;
        }),
      );

      const maxSeq = Math.max(maxSeqFromList, maxSeqFromArray);

      // Increment the max sequence number
      const newSeqNumber = (maxSeq + 1).toString().padStart(4, "0");

      // Create the document number based on details
      const newDocNumber = `${projectInfo.projectName}-${selectedArea.AreaCode || ""}-${
        selectedArea.AreaIdentificationCode || ""
      }-${selectedOrganization.OrganizationIdentificationCode || ""}-${
        selectedDocType.DocumentTypeCode || ""
      }-${newSeqNumber}`;

      // Update the deliverable row
      const updated = [...deliverables];
      updated[index] = {
        ...updated[index],
        seqNumber: newSeqNumber,
        docNumber: newDocNumber,
      };
      setDeliverables(updated);

      console.log(
        `Generated Seq: ${newSeqNumber}, Doc: ${newDocNumber} for row ${index}`,
      );
    } catch (error) {
      console.error("Error generating sequence and document number:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "generateSequenceAndDocNumber",
        "NewRequestWebPart",
      );
    }
  };

  // Used useEffect to fetch data on mount mean on loan only
  React.useEffect(() => {
    const initializeForm = async () => {
      try {
        // Load master data first
        await loadUsers();
        await fetchProjectTypes();
        await fetchDeliverables();
        await fetchArea();
        await fetchOrganization();
        await fetchDocumentType();
        // If initialData exists, populate form
        if (props.initialData) {
          setIsDraftMode(true);
          populateFormWithPayload(props.initialData);
        } else {
          await setCurrentUser();
        }
      } catch (error) {
        console.error("Error initializing form:", error);
        await ErrorLogger.logError(
          sp,
          error,
          "initializeForm",
          "NewRequestWebPart",
        );
      }
    };
    const fetchProjectTypes = async () => {
      try {
        const items = await sp.web.lists
          .getByTitle("ProjectTypeMasterList")
          .items.select("Id", "ProjectType", "IsActive")
          .filter("IsActive eq 'Yes'")();
        const options: IDropdownOption[] = items.map((item) => ({
          key: item.Id,
          text: item.ProjectType,
        }));
        const placeholder: IDropdownOption = {
          key: "",
          text: "Select Organization",
        };
        setProjectTypeOptions([placeholder, ...options]);
      } catch (error) {
        console.error("Error fetching project types:", error);
      }
    };

    const fetchDeliverables = async () => {
      try {
        const items = await sp.web.lists
          .getByTitle("DeliverablesMasterList")
          .items.select("Id", "Deliverables", "IsActive")
          .filter("IsActive eq 'Yes'")();
        const options: IDropdownOption[] = items.map((item) => ({
          key: item.Deliverables,
          text: item.Deliverables,
        }));
        const deliverablesData: DeliverablesMaster[] = items.map((item) => ({
          ID: item.Id,
          Deliverables: item.Deliverables,
        }));
        const placeholder: IDropdownOption = {
          key: "",
          text: "Select Deliverables",
        };
        setDeliverablesOptions([placeholder, ...options]);
        setDeliverablesMasterListArr(deliverablesData);
      } catch (error) {
        console.error("Error fetching deliverables:", error);
        await ErrorLogger.logError(
          sp,
          error,
          "fetchDeliverables",
          "NewRequestWebPart",
        );
      }
    };

    const fetchArea = async () => {
      try {
        const items = await sp.web.lists
          .getByTitle("AreaMasterList")
          .items.select(
            "Id",
            "Area",
            "AreaIdentificationCode",
            "AreaCode",
            "IsActive",
          )
          .filter("IsActive eq 'Yes'")();
        const options: IDropdownOption[] = items.map((item) => ({
          key: item.Area,
          text: item.Area,
        }));
        const areaData: AreaMaster[] = items.map((item) => ({
          ID: item.Id,
          Area: item.Area,
          AreaIdentificationCode: item.AreaIdentificationCode,
          AreaCode: item.AreaCode,
        }));
        const placeholder: IDropdownOption = { key: "", text: "Select Area" };
        setAreaOptions([placeholder, ...options]);
        setAreaMasterArr(areaData);
      } catch (error) {
        console.error("Error fetching area data:", error);
        await ErrorLogger.logError(sp, error, "fetchArea", "NewRequestWebPart");
      }
    };

    const fetchOrganization = async () => {
      try {
        const items = await sp.web.lists
          .getByTitle("OrganisationMasterList")
          .items.select(
            "Id",
            "Organisation",
            "OrganizationIdentificationCode",
            "IsActive",
          )
          .filter("IsActive eq 'Yes'")();
        const options: IDropdownOption[] = items.map((item) => ({
          key: item.Organisation,
          text: item.Organisation,
        }));
        const orgData: OrganizationMaster[] = items.map((item) => ({
          ID: item.Id,
          Organization: item.Organisation,
          OrganizationIdentificationCode: item.OrganizationIdentificationCode,
        }));
        const placeholder: IDropdownOption = {
          key: "",
          text: "Select Project Type",
        };
        setOrganizationOptions([placeholder, ...options]);
        setOrganizationMasterArr(orgData);
      } catch (error) {
        console.error("Error fetching Organization:", error);
        await ErrorLogger.logError(
          sp,
          error,
          "fetchOrganization",
          "NewRequestWebPart",
        );
      }
    };

    const fetchDocumentType = async () => {
      try {
        const items = await sp.web.lists
          .getByTitle("DocumentTypeMasterList")
          .items.select("Id", "DocumentType", "DocumentTypeCode", "IsActive")
          .filter("IsActive eq 'Yes'")();
        const options: IDropdownOption[] = items.map((item) => ({
          key: item.DocumentType,
          text: item.DocumentType,
        }));
        const docTypeData: DocumentTypeMaster[] = items.map((item) => ({
          ID: item.Id,
          DocumentType: item.DocumentType,
          DocumentTypeCode: item.DocumentTypeCode,
        }));
        const placeholder: IDropdownOption = {
          key: "",
          text: "Select Document Type",
        };
        setDocumentTypeOptions([placeholder, ...options]);
        setDocumentsTypeMasterArr(docTypeData);
      } catch (error) {
        console.error("Error fetching DocumentType:", error);
        await ErrorLogger.logError(
          sp,
          error,
          "fetchDocumentType",
          "NewRequestWebPart",
        );
      }
    };

    const loadUsers = async () => {
      try {
        const users = await sp.web.siteUsers();
        const options: UserOption[] = users.map((u) => ({
          value: u.Id,
          label: u.Title,
          email: u.Email,
        }));
        setUserOptions(options);
      } catch (error) {
        console.error("Error fetching users:", error);
        await ErrorLogger.logError(sp, error, "loadUsers", "NewRequestWebPart");
      }
    };

    const setCurrentUser = async () => {
      try {
        const currentUser = await sp.web.currentUser();
        const currentUserOption: UserOption = {
          value: currentUser.Id,
          label: currentUser.Title,
          email: currentUser.Email,
        };

        setSelectedUser(currentUserOption);
        handleProjectChange("preparedBy", currentUserOption.value);
      } catch (error) {
        console.error("Error fetching current user:", error);
        await ErrorLogger.logError(
          sp,
          error,
          "setCurrentUser",
          "NewRequestWebPart",
        );
      }
    };

    initializeForm();
    loadUsers();
    fetchProjectTypes();
    fetchDeliverables();
    fetchArea();
    fetchOrganization();
    fetchDocumentType();
    setCurrentUser();
  }, [props.initialData]);

  const populateFormWithPayload = (data: any) => {
    // Populate project info
    setProjectInfo({
      projectName: data.projectName || "",
      clientName: data.clientName || "",
      preparedBy: data.preparedById || 0,
      startDate: data.startDate || undefined,
      projectType: data.projectTypeId || "",
      overview: data.overview || "",
    });
    // Set selected user
    if (data.preparedById && userOptions.length > 0) {
      const user = userOptions.find((u) => u.value === data.preparedById);
      if (user) setSelectedUser(user);
    }
    // Populate deliverables from payload
    if (data.deliverables && data.deliverables.length > 0) {
      const mappedDeliverables = data.deliverables.map((d: any) => ({
        deliverable: d.deliverableName || d.deliverable || "",
        area: d.area || "",
        organisation: d.organisation || "",
        docType: d.docType || "",
        seqNumber: d.sequentialNumber || d.seqNumber || "",
        docNumber: d.docNumber || "",
        dueDate: d.dueDate ? new Date(d.dueDate) : undefined,
        assignedTo: d.assignedToId || 0,
      }));
      // Set deliverables after master data is loaded
      setTimeout(() => {
        setDeliverables(mappedDeliverables);
      }, 500);
    }
  };
  const handleUpdateExisting = async (action: "submit" | "draft") => {
    try {
      setShowSubmitLoader(true);
      const projectId = props.initialData.projectId;
      // Step 1: Update ProjectCreationList
      await sp.web.lists
        .getByTitle("ProjectCreationList")
        .items.getById(projectId)
        .update({
          ProjectName: projectInfo.projectName,
          ProjectOverview: projectInfo.overview,
          ClientName: projectInfo.clientName,
          ProjectStartDate: projectInfo.startDate,
          ProjectTypeId: parseInt(projectInfo.projectType) || 0,
          PreparedById: parseInt(projectInfo.preparedBy) || 0,
          Status: action === "submit" ? "Pending" : "Save as draft",
          SubmitStatus: action === "submit" ? "Yes" : "No",
        });
      // Step 2: Build new payload
      const projectTypeName =
        projectTypeOptions.find((opt) => opt.key === projectInfo.projectType)
          ?.text || "";
      const deliverablesPayload = deliverables.map((row, index) => {
        const selectedDeliverable = getDeliverableByText(row.deliverable);
        const selectedArea = getAreaByText(row.area);
        const selectedOrganization = getOrganizationByText(row.organisation);
        const selectedDocType = getDocumentTypeByText(row.docType);
        const assignedUserEmail =
          userOptions.find((u) => u.value === row.assignedTo)?.email || "";
        return {
          rowIndex: index + 1,
          deliverableId: selectedDeliverable?.ID || null,
          deliverableName: row.deliverable,
          areaId: selectedArea?.ID || null,
          area: row.area,
          organisationId: selectedOrganization?.ID || null,
          organisation: row.organisation,
          documentTypeId: selectedDocType?.ID || null,
          docType: row.docType,
          sequentialNumber: row.seqNumber,
          docNumber: row.docNumber,
          dueDate: row.dueDate ? row.dueDate.toISOString() : null,
          assignedToId: row.assignedTo,
          assignedToEmail: assignedUserEmail,
        };
      });
      // Build email payload (same as before)
      const today = new Date();
      const sentDate = today.toLocaleDateString("en-GB", {
        day: "numeric",
        month: "short",
        year: "numeric",
      });
      const regardsName = selectedUser?.label || "Management";
      const portalLink =
        "https://officeindia.sharepoint.com/sites/ESSA/SitePages/MyTasks.aspx";
      const uniqueAssigneeIds = [
        ...new Set(
          deliverables.filter((d) => d.assignedTo > 0).map((d) => d.assignedTo),
        ),
      ];
      const emailPayload = uniqueAssigneeIds.map((assigneeId) => {
        const assigneeUser = userOptions.find((u) => u.value === assigneeId);
        const tableRowsHtml = deliverables
          .map(
            (d, i) => `
        <tr>
          <td style="border:1px solid #ccc;padding:6px;text-align:center;">${i + 1}</td>
          <td style="border:1px solid #ccc;padding:6px;">${d.docNumber || "-"}</td>
          <td style="border:1px solid #ccc;padding:6px;">${d.deliverable || "-"}</td>
          <td style="border:1px solid #ccc;padding:6px;text-align:center;">0</td>
          <td style="border:1px solid #ccc;padding:6px;text-align:center;">-</td>
          <td style="border:1px solid #ccc;padding:6px;">TR-${String(i + 1).padStart(3, "0")}</td>
          <td style="border:1px solid #ccc;padding:6px;">${sentDate}</td>
          <td style="border:1px solid #ccc;padding:6px;">-</td>
        </tr>
      `,
          )
          .join("");
        const emailTable = `
        <table style="border-collapse:collapse;width:100%;max-width:900px;font-size:13px;margin-top:15px;">
          <thead>
            <tr style="background:#8eaada;font-weight:bold;">
              <th style="border:1px solid #ccc;padding:6px;text-align:center;">NO.#</th>
              <th style="border:1px solid #ccc;padding:6px;">Doc. #</th>
              <th style="border:1px solid #ccc;padding:6px;">Deliverables</th>
              <th style="border:1px solid #ccc;padding:6px;text-align:center;">Rev</th>
              <th style="border:1px solid #ccc;padding:6px;text-align:center;">Status</th>
              <th style="border:1px solid #ccc;padding:6px;">Transmittal</th>
              <th style="border:1px solid #ccc;padding:6px;">Sent Dated</th>
              <th style="border:1px solid #ccc;padding:6px;">Remarks</th>
            </tr>
          </thead>
          <tbody>
            ${tableRowsHtml}
          </tbody>
        </table>
      `;
        const emailBody = `
        Dear ${assigneeUser?.label || "Vendor"},<br/><br/>
        A new task has been assigned to you for the project
        <strong>"${projectInfo.projectName}"</strong>.<br/><br/>
        Please review the task details below:<br/>
        ${emailTable}
        <br/><br/>
        <a href="${portalLink}" style="color:#0078d4;font-weight:bold;text-decoration:underline;">Click here to view in portal</a>
        <br/><br/>
        Regards,<br/>
        ${regardsName}
      `;
        return {
          toUserId: assigneeId,
          toUserEmail: assigneeUser?.email || "",
          ccUserId: selectedUser?.value || null,
          ccUserEmail: selectedUser?.email || "",
          assigneeName: assigneeUser?.label || "Vendor",
          subject: `New Task Assigned – ${projectInfo.projectName}`,
          body: emailBody,
        };
      });
      // Step 3: Build final payload
      const payload = {
        projectCreationId: projectId,
        projectName: projectInfo.projectName,
        projectType: projectTypeName,
        Status: action === "submit" ? "Pending" : "Save as draft",
        SubmitStatus: action === "submit" ? "Yes" : "No",
        preparedByName: selectedUser?.label || "",
        preparedById: selectedUser?.value || null,
        preparedByEmail: selectedUser?.email || "",
        deliverables: deliverablesPayload,
        emails: emailPayload,
      };
      // Step 4: Update RequestProcessingQueue
      const queueItems = await sp.web.lists
        .getByTitle("RequestProcessingQueue")
        .items.select("ID")
        .filter(`ProjectCreationID eq ${projectId}`)();
      if (queueItems.length > 0) {
        await sp.web.lists
          .getByTitle("RequestProcessingQueue")
          .items.getById(queueItems[0].ID)
          .update({
            Status: action === "submit" ? "Pending" : "Save as draft",
            ProcessingStatus: "Pending",
            PayloadJSON: JSON.stringify(payload),
          });
      } else {
        // If no queue item exists, create one
        await sp.web.lists.getByTitle("RequestProcessingQueue").items.add({
          Title: projectInfo.projectName,
          ProjectCreationID: projectId,
          PayloadJSON: JSON.stringify(payload),
          Status: action === "submit" ? "Pending" : "Save as draft",
          ProcessingStatus: "Pending",
        });
      }
      setShowSubmitLoader(false);
      setPopupConfig({
        type: "success",
        title: "Success",
        message:
          action === "submit"
            ? "Request submitted successfully!"
            : "Draft saved successfully!",
      });
      setShowPopup(true);
    } catch (error) {
      console.error("Error updating draft:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "handleUpdateExisting",
        "NewRequestWebPart",
      );
      setShowSubmitLoader(false);
      setPopupConfig({
        type: "error",
        title: "Error",
        message: "Failed to update. Please try again.",
      });
      setShowPopup(true);
    }
  };

  const fetchDeliverablesByProjectType = async (projectType: string) => {
    try {
      const items = await sp.web.lists
        .getByTitle("ProjectTypeDeliverablesMaster")
        .items.select(
          "Deliverables/Deliverables",
          "Deliverables/Id",
          "Area/Area",
          "Area/Id",
          "Organization/Organisation",
          "Organization/Id",
          "DocumentType/DocumentType",
          "DocumentType/Id",
        )
        .expand("Deliverables", "Area", "Organization", "DocumentType")
        .filter(`ProjectType/Id eq '${projectType}'`)
        .top(4999)();
      if (items.length > 0) {
        const mapped: DeliverableItem[] = await Promise.all(
          items.map(async (item, index) => {
            const deliverableItem: DeliverableItem = {
              deliverable: item.Deliverables?.Deliverables || "",
              area: item.Area?.Area || "",
              organisation: item.Organization?.Organisation || "",
              docType: item.DocumentType?.DocumentType || "",
              seqNumber: "",
              docNumber: "",
              dueDate: undefined,
              assignedTo: 0,
            };

            // Generate sequence and doc number for each item
            const tempDeliverables = [deliverableItem];
            const currentIndex = 0;

            const selectedArea = getAreaByText(deliverableItem.area);
            const selectedOrganization = getOrganizationByText(
              deliverableItem.organisation,
            );
            const selectedDocType = getDocumentTypeByText(
              deliverableItem.docType,
            );

            if (
              projectInfo.projectName &&
              selectedArea &&
              selectedOrganization &&
              selectedDocType
            ) {
              try {
                const existingItems = await sp.web.lists
                  .getByTitle("DeliverablesDetails")
                  .items.select(
                    "SequenticalNumber",
                    "ProjectCreationListID/ID",
                    "ProjectCreationListID/ProjectName",
                  )
                  .expand("ProjectCreationListID")
                  .filter(
                    `ProjectCreationListID/ProjectName eq '${projectInfo.projectName}' 
                                        and Area eq '${selectedArea.Area}' 
                                        and Organisation eq '${selectedOrganization.Organization}' 
                                        and DocumentType eq '${selectedDocType.DocumentType}'`,
                  )
                  .orderBy("ID", true)
                  .top(4999)();

                const maxSeqFromList =
                  existingItems.length > 0
                    ? Math.max(
                        ...existingItems.map((r) =>
                          r.SequenticalNumber
                            ? parseInt(r.SequenticalNumber)
                            : 0,
                        ),
                      )
                    : 0;

                const maxSeqFromArray = Math.max(
                  ...tempDeliverables.map((r, i) => {
                    if (i === currentIndex) return 0;
                    const isMatch =
                      r.area === deliverableItem.area &&
                      r.organisation === deliverableItem.organisation &&
                      r.docType === deliverableItem.docType;
                    return isMatch && r.seqNumber ? parseInt(r.seqNumber) : 0;
                  }),
                );

                const maxSeq = Math.max(maxSeqFromList, maxSeqFromArray);
                const newSeqNumber = (maxSeq + 1).toString().padStart(4, "0");

                deliverableItem.seqNumber = newSeqNumber;
                deliverableItem.docNumber = `${projectInfo.projectName}-${selectedArea.AreaCode}-${
                  selectedArea.AreaIdentificationCode
                }-${selectedOrganization.OrganizationIdentificationCode}-${
                  selectedDocType.DocumentTypeCode
                }-${newSeqNumber}`;
              } catch (error) {
                console.error(
                  "Error generating sequence for pre-filled deliverables:",
                  error,
                );
              }
            }

            return deliverableItem;
          }),
        );

        const normalized = normalizeSequenceNumbers(mapped);
        setDeliverables(normalized);
      } else {
        setDeliverables([
          {
            deliverable: "",
            area: "",
            organisation: "",
            docType: "",
            seqNumber: "",
            docNumber: "",
            dueDate: undefined,
            assignedTo: 0,
          },
        ]);
      }
    } catch (error) {
      console.error("Error fetching ProjectTypeDeliverablesMaster:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "fetchDeliverablesByProjectType",
        "NewRequestWebPart",
      );
      setDeliverables([
        {
          deliverable: "",
          area: "",
          organisation: "",
          docType: "",
          seqNumber: "",
          docNumber: "",
          dueDate: undefined,
          assignedTo: 0,
        },
      ]);
    }
  };

  const handleCancelPopup = () => {
    setShowPopup(false);
    if (popupConfig.type === "confirmation") {
      setShowSubmitLoader(false);
    }
  };

  const handleCloseValidation = () => {
    setShowPopup(false);
    setShowSubmitLoader(false);
  };

  // Updated handleSubmit function
  const handleSubmit = (action: "submit" | "draft") => {
    // Check if editing existing draft or new request
  const isEditingDraft = props.initialData && props.initialData.projectId;
  
if (isEditingDraft) {

  setSubmitAction(action);
  setShowSubmitLoader(true);

  if (!validateForm(action)) {
    return;
  }

  setPopupConfig({
    type: "confirmation",
    title:
      action === "submit"
        ? "Confirm Submission"
        : "Confirm Save as Draft",
    message:
      action === "submit"
        ? "Are you sure you want to submit this request?"
        : "Are you sure you want to save this request as draft?",
  });

  setShowPopup(true);
  setShowSubmitLoader(false);

  return;
} else {
    // New request flow
  setSubmitAction(action);
  setShowSubmitLoader(true);

 if (!validateForm(action)) {
    return;
  }

  setPopupConfig({
    type: "confirmation",
    title: action === "submit" ? "Confirm Submission" : "Confirm Save as Draft",
    message: action === "submit" 
      ? "Are you sure you want to submit this request?" 
      : "Are you sure you want to save this request as draft?",
  });
  setShowPopup(true);
  setShowSubmitLoader(false);
   }
};

//    const handleSuccessOk = () => {
//   setShowPopup(false);
//   if (isDraftMode) {
//   if (props.onCancel) props.onCancel();
// } else{
//      window.location.reload();
//   }
// };

const handleSuccessOk = () => {

  setShowPopup(false);

  if (props.onCancel) {
    props.onCancel();
  }

  window.location.reload();
};

  const handleConfirm = async () => {
    
    setShowPopup(false);
    setShowSubmitLoader(true);

    const isEditingDraft = props.initialData && props.initialData.projectId;

if (isEditingDraft) {
  await handleUpdateExisting(submitAction);
  return;
}

    try {
      // Step 1: Create the project item in ProjectCreationList
     const projectRequestItem = await sp.web.lists
  .getByTitle("ProjectCreationList")
  .items.add({
    Title: projectInfo.projectName,
    ProjectName: projectInfo.projectName,
    ProjectOverview: projectInfo.overview,
    Status: submitAction === "submit" ? "Pending" : "Save as draft",
    SubmitStatus: submitAction === "submit" ? "Yes" : "No",
    ProjectTypeId: parseInt(projectInfo.projectType) || 0,
    ClientName: projectInfo.clientName,
    ProjectStartDate: projectInfo.startDate,
    PreparedById: parseInt(projectInfo.preparedBy) || 0,
  });

      const projectCreationListId = projectRequestItem.Id;
      console.log("Created Project ID:", projectCreationListId);

      // Step 2: Build deliverables payload
      const deliverablesPayload = deliverables.map((row, index) => {
        const selectedDeliverable = getDeliverableByText(row.deliverable);
        const selectedArea = getAreaByText(row.area);
        const selectedOrganization = getOrganizationByText(row.organisation);
        const selectedDocType = getDocumentTypeByText(row.docType);

        const assignedUserEmail =
          userOptions.find((u) => u.value === row.assignedTo)?.email ?? "";

        return {
          rowIndex: index + 1,
          deliverableId: selectedDeliverable?.ID ?? null,
          deliverableName: row.deliverable,
          areaId: selectedArea?.ID ?? null,
          area: row.area,
          organisationId: selectedOrganization?.ID ?? null,
          organisation: row.organisation,
          documentTypeId: selectedDocType?.ID ?? null,
           docType: row.docType,
          sequentialNumber: row.seqNumber,
          docNumber: row.docNumber,
          dueDate: row.dueDate ? row.dueDate.toISOString() : null,
          assignedToId: row.assignedTo,
          assignedToEmail: assignedUserEmail,
          
        };
      });

      // Step 3: Build email payload with full HTML body
      const today = new Date();
      const sentDate = today.toLocaleDateString("en-GB", {
        day: "numeric",
        month: "short",
        year: "numeric",
      });

      const regardsName = selectedUser?.label ?? "Management";
      const portalLink =
        "https://officeindia.sharepoint.com/sites/ESSA/SitePages/MyTasks.aspx";

      // Get unique assignees — one email per person
      const uniqueAssigneeIds = [
        ...new Set(
          deliverables.filter((d) => d.assignedTo > 0).map((d) => d.assignedTo),
        ),
      ];

      const emailPayload = uniqueAssigneeIds.map((assigneeId) => {
        const assigneeUser = userOptions.find((u) => u.value === assigneeId);

        // Build table rows HTML
        const tableRowsHtml = deliverables
          .map(
            (d, i) => `
        <tr>
          <td style="border:1px solid #ccc;padding:6px;text-align:center;">${
            i + 1
          }</td>
          <td style="border:1px solid #ccc;padding:6px;">${
            d.docNumber ?? "-"
          }</td>
          <td style="border:1px solid #ccc;padding:6px;">${
            d.deliverable ?? "-"
          }</td>
          <td style="border:1px solid #ccc;padding:6px;text-align:center;">0</td>
          <td style="border:1px solid #ccc;padding:6px;text-align:center;">-</td>
          <td style="border:1px solid #ccc;padding:6px;">TR-${String(
            i + 1,
          ).padStart(3, "0")}</td>
          <td style="border:1px solid #ccc;padding:6px;">${sentDate}</td>
          <td style="border:1px solid #ccc;padding:6px;">-</td>
        </tr>`,
          )
          .join("");

        // Build full email table HTML
        const emailTable = `
        <table style="border-collapse:collapse;width:100%;max-width:900px;font-size:13px;margin-top:15px;">
          <thead>
            <tr style="background:#8eaada;font-weight:bold;">
              <th style="border:1px solid #ccc;padding:6px;text-align:center;">NO.#</th>
              <th style="border:1px solid #ccc;padding:6px;">Doc. #</th>
              <th style="border:1px solid #ccc;padding:6px;">Deliverables</th>
              <th style="border:1px solid #ccc;padding:6px;text-align:center;">Rev</th>
              <th style="border:1px solid #ccc;padding:6px;text-align:center;">Status</th>
              <th style="border:1px solid #ccc;padding:6px;">Transmittal</th>
              <th style="border:1px solid #ccc;padding:6px;">Sent Dated</th>
              <th style="border:1px solid #ccc;padding:6px;">Remarks</th>
            </tr>
          </thead>
          <tbody>
            ${tableRowsHtml}
          </tbody>
        </table>`;

        // Build full email body HTML
        const emailBody = `
        Dear ${assigneeUser?.label ?? "Vendor"},<br/><br/>
        A new task has been assigned to you for the project 
        <strong>"${projectInfo.projectName}"</strong>.<br/><br/>
        Please review the task details below:<br/>
        ${emailTable}
        <br/><br/>
        <a 
          href="${portalLink}" 
          style="color:#0078d4;font-weight:bold;text-decoration:underline;"
        >
          Click here to view in portal
        </a>
        <br/><br/>
        Regards,<br/>
        ${regardsName}`;

        // Build tableRows array (for JSON format compatibility)
        const tableRows = deliverables.map((d, i) => ({
          no: i + 1,
          docNumber: d.docNumber ?? "-",
          deliverable: d.deliverable ?? "-",
          revision: "0",
          status: "-",
          transmittal: `TR-${String(i + 1).padStart(3, "0")}`,
          sentDate: sentDate,
          remarks: "-",
        }));

        return {
          toUserId: assigneeId,
          toUserEmail: assigneeUser?.email ?? "",
          ccUserId: selectedUser?.value ?? null,
          ccUserEmail: selectedUser?.email ?? "",
          assigneeName: assigneeUser?.label ?? "Vendor",
          subject: `New Task Assigned – ${projectInfo.projectName}`,
          body: emailBody,
          tableRows: tableRows,
        };
      });

      const projectTypeName =
        projectTypeOptions.find((opt) => opt.key === projectInfo.projectType)
          ?.text ?? "";

      // Step 4: Build final payload
      const payload = {
  projectCreationId: projectCreationListId,
  projectName: projectInfo.projectName,
  projectType: projectTypeName,
  Status: submitAction === "submit" ? "Pending" : "Save as draft",
  SubmitStatus: submitAction === "submit" ? "Yes" : "No",
  preparedByName: selectedUser?.label ?? "",
  preparedById: selectedUser?.value ?? null,
  preparedByEmail: selectedUser?.email ?? "",
  deliverables: deliverablesPayload,
  emails: emailPayload,
  
};

      console.log("Payload to queue:", JSON.stringify(payload, null, 2));

      // Step 5: Drop one item into RequestProcessingQueue
      await sp.web.lists.getByTitle("RequestProcessingQueue").items.add({
  Title: projectInfo.projectName,
  ProjectCreationID: projectCreationListId,
  PayloadJSON: JSON.stringify(payload),
  Status: submitAction === "submit" ? "Pending" : "Save as draft",
  ProcessingStatus: "Pending",
});

      // Step 6: Show success immediately — Power Automate handles the rest
      setShowSubmitLoader(false);
      setPopupConfig({
        type: "success",
        title: "Success",
        message:submitAction === "submit"
      ? "Request submitted successfully!"
      : "Request saved successfully!",
      });
      setShowPopup(true);
    } catch (error) {
      console.error("Error submitting request:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "handleConfirm",
        "NewRequestWebPart",
      );
      setShowSubmitLoader(false);
      setPopupConfig({
        type: "error",
        title: "Error",
        message: "Failed to submit request. Please try again.",
      });
      setShowPopup(true);
    }
  };

  const handleProjectChange = (field: string, value: any) => {
    setProjectInfo((prev) => ({ ...prev, [field]: value }));
    console.log(projectInfo);
  };

  type DeliverableField = keyof DeliverableItem;
  const handleDeliverableChange = async (
    index: number,
    field: DeliverableField,
    value: string | number | Date | null | undefined,
  ) => {
    const updated = [...deliverables];

    if (field === "dueDate" && value === null) {
      updated[index][field] = undefined as never;
    } else {
      updated[index][field] = value as never;
    }

    setDeliverables(updated);

    // Generate sequence and document number when area, organization, or document type changes
    if (field === "area" || field === "organisation" || field === "docType") {
      if (field === "area" || field === "organisation") {
        // For area and organization changes, use 'Project Type' as source
        await generateSequenceAndDocNumber(index, "Project Type");
      } else if (field === "docType") {
        // For document type changes, use 'Document Type' as source
        await generateSequenceAndDocNumber(index, "Document Type");
      }
    }
  };

  const addDeliverableRow = () => {
    setDeliverables((prev) => [
      ...prev,
      {
        deliverable: "",
        area: "",
        organisation: "",
        docType: "",
        seqNumber: "",
        docNumber: "",
        dueDate: undefined,
        assignedTo: 0,
      },
    ]);
  };

  const deleteDeliverableRow = (index: number) => {
    if (deliverables.length === 1) {
      setPopupConfig({
        type: "validation",
        title: "Validation Error",
        message: "You can't delete last row.",
      });
      setShowPopup(true);
      return;
    }

    const updated = [...deliverables];
    updated.splice(index, 1);
    setDeliverables(updated);
  };

  const handleCancel = () => {
    if (props.onCancel) {
      props.onCancel();
    } else {
      // history.push("/Dashboard.aspx");
    }
  };

  const normalizeSequenceNumbers = (rows: DeliverableItem[]) => {
    const counter: Record<string, number> = {};

    return rows.map((row) => {
      const key = `${row.area}|${row.organisation}|${row.docType}`;

      if (!counter[key]) {
        counter[key] = parseInt(row.seqNumber || "0", 10) || 1;
      } else {
        counter[key] += 1;
      }

      const seq = counter[key].toString().padStart(4, "0");

      return {
        ...row,
        seqNumber: seq,
        docNumber: row.docNumber.replace(/\d{4}$/, seq),
      };
    });
  };

  return (
    <div className={styles.newRequest}>
      {/* 🔄 SUBMIT TASK LOADER */}

      {showSubmitLoader && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            width: "100%",
            height: "100%",
            backgroundColor: "rgba(0,0,0,0.45)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            zIndex: 9999,
          }}
        >
          <div className="text-center">
            <img
              src={require("../../../CustomAsset/birdloader.gif")}
              alt="Submitting..."
              style={{ width: "90px", height: "90px" }}
            />
            <p className="text-white mt-3 fw-semibold">
              Submitting, please wait...
            </p>
          </div>
        </div>
      )}
      <h2>New Request</h2>

      {/* Project Information Section */}
      <div className={styles.section}>
        <h3>Project Information</h3>

        <div className={styles.formRow}>
          <TextField
            label="Project Name"
            placeholder="Enter Project Name"
            required
            value={projectInfo.projectName}
            onChange={(_, val) => {
              handleProjectChange("projectName", val);
              setFieldErrors((prev) => ({ ...prev, projectName: false }));
            }}
            className={fieldErrors.projectName ? styles.inputError : ""}
          />
          <TextField
            label="Client Name"
            placeholder="Enter Client Name"
            required
            value={projectInfo.clientName}
            onChange={(_, val) => {
              handleProjectChange("clientName", val);
              setFieldErrors((prev) => ({ ...prev, clientName: false }));
            }}
            className={fieldErrors.clientName ? styles.inputError : ""}
          />
        </div>

        <div className={styles.formRow}>
          <div className={styles.selectContainer}>
            <label className={styles.selectLabel}>Prepared By</label>
            <Select
              options={userOptions}
              value={selectedUser}
              onChange={(option: UserOption | null) => {
                setSelectedUser(option);
                handleProjectChange("preparedBy", option?.value || 0);
                setFieldErrors((prev) => ({ ...prev, preparedBy: false }));
              }}
              placeholder="Select a user..."
              className={fieldErrors.preparedBy ? styles.inputError : ""}
              classNamePrefix="react-select"
            />
          </div>
          <div>
            <label style={{ fontWeight: 600 }}>
              Project Start Date <span style={{ color: "red" }}>*</span>
            </label>

            <DatePicker
              placeholder="Enter Project Start Date"
              value={projectInfo.startDate}
              onSelectDate={(date) => {
                handleProjectChange("startDate", date);
                setFieldErrors((prev) => ({ ...prev, startDate: false }));
              }}
              formatDate={(date) =>
                date ? date.toLocaleDateString("en-GB") : ""
              }
              className={fieldErrors.startDate ? styles.inputError : ""}
            />
          </div>
        </div>

        <div className={styles.formRow}>
          <Dropdown
            label="Project Type"
            required
            options={projectTypeOptions}
            selectedKey={projectInfo.projectType}
            onChange={async (_, option: IDropdownOption | undefined) => {
              if (
                !projectInfo.projectName ||
                projectInfo.projectName.trim() === ""
              ) {
                setPopupConfig({
                  type: "validation",
                  title: "Validation Error",
                  message:
                    "⚠️ Please enter Project Name before selecting Project Type.",
                });
                setShowPopup(true);
                return;
              }
              const selectedType = option?.key as string;
              handleProjectChange("projectType", selectedType);
              setFieldErrors((prev) => ({ ...prev, projectType: false }));
              await fetchDeliverablesByProjectType(selectedType);
            }}
            className={fieldErrors.projectType ? styles.projectTypeError : ""}
          />
          <TextField
            label="Project Overview"
            placeholder="Enter Project Overview"
            required
            multiline
            rows={3}
            value={projectInfo.overview}
            onChange={(_, val) => {
              handleProjectChange("overview", val);
              setFieldErrors((prev) => ({ ...prev, overview: false }));
            }}
            className={fieldErrors.overview ? styles.inputError : ""}
          />
        </div>
      </div>

      {/* Deliverables Section */}
      <div className="d-flex align-items-center  justify-content-between mb-2">
        <h4
          style={{ margin: "0px" }}
          className="font-16 fw-bold text-dark mb-0"
        >
          Deliverables
        </h4>
        <IconButton
          iconProps={{ iconName: "Add" }}
          title="Add Row"
          ariaLabel="Add Row"
          onClick={addDeliverableRow}
          className={styles.addButton}
        />
      </div>
      <div className={styles.tableWrapper}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th style={{ minWidth: "70px" }}>SNo</th>
              <th>Deliverables*</th>
              <th>Area*</th>
              <th>Organisation*</th>
              <th>Document Type*</th>
              <th>Seq. No</th>
              <th className={styles.wrapText}>Doc. No</th>
              <th>Due Date*</th>
              <th>Assigned To*</th>
              <th style={{ minWidth: "80px" }}>Action</th>
            </tr>
          </thead>
          <tbody>
            {deliverables.map((row, index) => (
              <tr key={index}>
                <td style={{ minWidth: "70px" }}>{index + 1}</td>

                {/* Deliverable Column */}
                <td title={row.deliverable || "Select a deliverable"}>
                  <Dropdown
                    options={DeliverablesOptions}
                    selectedKey={row.deliverable}
                    onChange={(_: any, option: IDropdownOption | undefined) => {
                      handleDeliverableChange(
                        index,
                        "deliverable",
                        option?.key,
                      );
                      setFieldErrors((prev) => ({
                        ...prev,
                        [`deliverable_${index}`]: false,
                        minimumDeliverable: false,
                      }));
                    }}
                    placeholder="Select Deliverable"
                    className={
                      fieldErrors[`deliverable_${index}`]
                        ? styles.tableDropdownError
                        : ""
                    }
                  />
                </td>

                <td title={row.area || "Select an area"}>
                  <Dropdown
                    options={AreaOptions}
                    selectedKey={row.area}
                    onChange={(_: any, option: IDropdownOption | undefined) => {
                      handleDeliverableChange(index, "area", option?.key);
                      setFieldErrors((prev) => ({
                        ...prev,
                        [`area_${index}`]: false,
                      }));
                    }}
                    placeholder="Select Area"
                    className={
                      fieldErrors[`area_${index}`]
                        ? styles.tableDropdownError
                        : ""
                    }
                  />
                </td>

                <td title={row.organisation || "Select an organisation"}>
                  <Dropdown
                    options={OrganizationOptions}
                    selectedKey={row.organisation}
                    onChange={(_: any, option: IDropdownOption | undefined) => {
                      handleDeliverableChange(
                        index,
                        "organisation",
                        option?.key,
                      );
                      setFieldErrors((prev) => ({
                        ...prev,
                        [`organisation_${index}`]: false,
                      }));
                    }}
                    placeholder="Select Organisation"
                    className={
                      fieldErrors[`organisation_${index}`]
                        ? styles.tableDropdownError
                        : ""
                    }
                  />
                </td>

                <td title={row.docType || "Select a document type"}>
                  <Dropdown
                    options={DocumentTypeOptions}
                    selectedKey={row.docType}
                    onChange={async (
                      _: any,
                      option: IDropdownOption | undefined,
                    ) => {
                      const selectedDocumentType = option?.key as string;
                      await handleDeliverableChange(
                        index,
                        "docType",
                        selectedDocumentType,
                      );
                      setFieldErrors((prev) => ({
                        ...prev,
                        [`docType_${index}`]: false,
                      }));
                    }}
                    placeholder="Select Document Type"
                    className={
                      fieldErrors[`docType_${index}`]
                        ? styles.tableDropdownError
                        : ""
                    }
                  />
                </td>

                {/* Seq. No Column */}
                <td title={row.seqNumber || "Sequence number"}>
                  <TextField value={row.seqNumber} disabled />
                </td>

                {/* Doc. No Column */}
                <td title={row.docNumber || "Full document number"}>
                  <TextField value={row.docNumber} disabled />
                </td>

                {/* Due Date Column - MANDATORY */}

                <td>
                  <DatePicker
                    value={row.dueDate}
                    onSelectDate={(date) => {
                      handleDeliverableChange(index, "dueDate", date);

                      setFieldErrors((prev) => ({
                        ...prev,
                        [`dueDate_${index}`]: false,
                      }));
                    }}
                    formatDate={(date) =>
                      date ? date.toLocaleDateString("en-GB") : ""
                    }
                    className={
                      fieldErrors[`dueDate_${index}`]
                        ? styles.tableDateError
                        : ""
                    }
                  />
                </td>

                {/* Assigned To Column - MANDATORY */}

                <td>
                  <div className={styles.selectContainer}>
                    <Select
                      options={userOptions}
                      value={findUserOption(userOptions, row.assignedTo)}
                      onChange={(option: UserOption | null) => {
                        handleDeliverableChange(
                          index,
                          "assignedTo",
                          option?.value || 0,
                        );

                        setFieldErrors((prev) => ({
                          ...prev,
                          [`assignedTo_${index}`]: false,
                        }));
                      }}
                      classNamePrefix="react-select"
                      className={`${styles.customSelect} ${
                        fieldErrors[`assignedTo_${index}`]
                          ? styles.tableSelectError
                          : ""
                      }`}
                      menuPortalTarget={document.body}
                      menuPosition="fixed"
                      styles={{
                        menuPortal: (base: any) => ({
                          ...base,
                          zIndex: 9999,
                        }),
                      }}
                    />
                  </div>
                </td>

                {/* Action Column */}
                <td style={{ minWidth: "80px" }}>
                  <IconButton
                    iconProps={{ iconName: "Delete" }}
                    onClick={() => deleteDeliverableRow(index)}
                    className={styles.deleteButton}
                  />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className={styles.buttonRow}>
  <PrimaryButton
    style={{ height: "36px" }}
    text="Save as Draft"
    iconProps={{ iconName: "Save" }}
    onClick={() => handleSubmit("draft")}
    disabled={showSubmitLoader}
    className={styles.saveDraftButton}
  />
  
  <PrimaryButton
    style={{ height: "36px" }}
    text="Submit"
    iconProps={{ iconName: "CheckMark" }}
    onClick={() => handleSubmit("submit")}
    disabled={showSubmitLoader}
    className={styles.submitButton}
  />

  <DefaultButton
    style={{ height: "36px" }}
    text="Cancel"
    iconProps={{ iconName: "Cancel" }}
    onClick={handleCancel}
    className={styles.cancelButton}
  />
</div>
      {/* Custom Popup */}
      <CustomPopup
        isOpen={showPopup}
        type={popupConfig.type}
        title={popupConfig.title}
        message={popupConfig.message}
        onConfirm={handleConfirm}
        onCancel={handleCancelPopup}
        onClose={handleCloseValidation}
        onSuccessOk={handleSuccessOk}
      />
    </div>
  );
};

export default NewRequest;
