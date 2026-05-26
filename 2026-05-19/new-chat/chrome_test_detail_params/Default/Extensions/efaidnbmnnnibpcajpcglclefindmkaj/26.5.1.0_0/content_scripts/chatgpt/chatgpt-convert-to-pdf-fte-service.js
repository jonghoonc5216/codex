/*************************************************************************
* ADOBE CONFIDENTIAL
* ___________________
*
*  Copyright 2015 Adobe Systems Incorporated
*  All Rights Reserved.
*
* NOTICE:  All information contained herein is, and remains
* the property of Adobe Systems Incorporated and its suppliers,
* if any.  The intellectual and technical concepts contained
* herein are proprietary to Adobe Systems Incorporated and its
* suppliers and are protected by all applicable intellectual property laws,
* including trade secret and or copyright laws.
* Dissemination of this information or reproduction of this material
* is strictly forbidden unless prior written permission is obtained
* from Adobe Systems Incorporated.
**************************************************************************/
import{sendErrorLog,sendAnalytics}from"../utils/util.js";import{createFteTooltip,updateFteToolTipCoolDown,acrobatTouchPointClicked}from"../utils/fte-utils.js";import state from"./state.js";const FTE_TOOLTIP_CONTAINER_CLASS="acrobat-fte-tooltip-container",FTE_STYLESHEET_MARKER="data-acrobat-deepresearch-fte-css",documentListenerOptions={signal:state?.eventControllerSignal};export const removeFteTooltip=()=>{const e=document.getElementsByClassName(FTE_TOOLTIP_CONTAINER_CLASS);if(e.length>0){const t=e[0];t.clickOutsideHandler&&document.removeEventListener("click",t.clickOutsideHandler,documentListenerOptions),t.parentElement?.classList.remove("acrobat-fte-active"),t.remove()}document.querySelector(`link[${FTE_STYLESHEET_MARKER}]`)?.remove()};export const attachConvertToPdfFteListeners=(e,{source:t,workflow:o})=>{const r=()=>{removeFteTooltip()};e.clickOutsideHandler=n=>{e.contains(n.target)||(r(),sendAnalytics([["DCBrowserExt:DirectVerb:Fte:Dismissed",{source:t,workflow:o}]]))};const n=e.querySelector(".acrobat-fte-tooltip-button");n&&n.addEventListener("click",e=>{e.stopPropagation(),r(),sendAnalytics([["DCBrowserExt:DirectVerb:Fte:Closed",{source:t,workflow:o}]])})};export const addFte=async({touchpointClass:e,fteType:t,storageKey:o,source:r,workflow:n})=>{try{const s=await chrome.runtime.sendMessage({main_op:"chatgpt-convert-to-pdf-init"}),c=s?.deepResearch;if(!c?.enableFte)return;const i=c?.fteConfig,a=document.getElementsByClassName(e)[0];if(!a)return;if(document.getElementsByClassName(FTE_TOOLTIP_CONTAINER_CLASS).length>0)return;if(document.head&&!document.querySelector(`link[${FTE_STYLESHEET_MARKER}]`)){const e=document.createElement("link");e.rel="stylesheet",e.href=chrome.runtime.getURL("browser/css/fte.css"),e.setAttribute(FTE_STYLESHEET_MARKER,"1"),document.head.insertBefore(e,document.head.firstChild)}const l={title:chrome.i18n.getMessage("chatgptConvertToPDFFteTitle"),description:chrome.i18n.getMessage("chatgptConvertToPDFFteDescription"),button:chrome.i18n.getMessage("closeButton")},d=createFteTooltip(l,t);attachConvertToPdfFteListeners(d,{source:r,workflow:n}),document.addEventListener("click",d.clickOutsideHandler,{once:!0,...documentListenerOptions}),a.classList.add("acrobat-fte-active"),a.appendChild(d),sendAnalytics([["DCBrowserExt:DirectVerb:Fte:Shown",{source:r,workflow:n}]]),await updateFteToolTipCoolDown(i,o)}catch{sendErrorLog(`ConvertToPdfFte:${t}`,`Failure adding FTE tooltip for ${r}`)}};export const markConsumed=async e=>{removeFteTooltip(),await acrobatTouchPointClicked(e)};