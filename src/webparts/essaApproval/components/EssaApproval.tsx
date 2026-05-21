import * as React from 'react';
import { useState, useEffect, useRef } from "react";
import Select from "react-select";
import styles from './EssaApproval.module.scss';
import { IEssaApprovalProps } from './IEssaApprovalProps';
import "@pnp/sp/attachments";

// --- PnP & Context Imports ---
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import Provider from "../../../GlobalContext/provider";

// --- Navigation Components ---
import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";
import "../../verticalSideBar/components/VerticalSidebar.scss";

const contentWrapStyle: React.CSSProperties = {
  padding: '8px 16px 16px',
  minWidth: 0
};

const EssaApproval: React.FC<IEssaApprovalProps> = (props) => {
  const { hasTeamsContext, userDisplayName, context, siteUrl } = props;
  const elementRef = useRef<HTMLDivElement>(null);
   // --- Initialize PnP JS ---
  const sp: SPFI = getSP(context);
  const renderContent = () => {
    return (
    <section className={styles.essaApproval}>
      <h1>Hello</h1>
    </section>
    );
  };


  return (
    <Provider>
      <div id="wrapper" ref={elementRef}>
        <div className="app-menu" id="myHeader">
          <VerticalSideBar _context={sp} />
        </div>
        <div className="content-page">
          <HorizontalNavbar _context={sp} siteUrl={siteUrl} context={context} />
          <div style={contentWrapStyle}>{renderContent()}</div>
        </div>
      </div>
    </Provider>
  );
};

export default EssaApproval;
