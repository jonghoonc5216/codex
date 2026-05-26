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
import{useCallback}from"react";import{sendAnalyticsEvent}from"../utils/fabUtils";export const useFABPositioning=({containerRef:t,buttonRef:e,fabManager:o})=>{const n=useCallback(e=>{const o=t.current;o&&e&&(o.style.top=`${e}px`,o.style.bottom="auto")},[t]);return{rePositionFABIfOverlapping:useCallback(()=>{try{const t=e.current;if(!t)return;if(void 0!==o.fabAutoRepositioningTop)return void n(o.fabAutoRepositioningTop);let i=0,r=null;for(;i<5;){const e="function"==typeof getClickableOverlappingElement?getClickableOverlappingElement(t,r):[];if(!e?.length)break;let o=Number.MAX_SAFE_INTEGER;if(e.forEach(t=>{const e=t.getBoundingClientRect();e.top<o&&(o=e.top)}),o===Number.MAX_SAFE_INTEGER)break;const n=t.getBoundingClientRect(),a=n.height,s=o-a,{left:l,right:p,bottom:g,width:c,height:u,x:f,y:b}=n;r={top:s,left:l,right:p,bottom:g,width:c,height:u,x:f,y:b},i++}if(o.fabAutoRepositioningTop=r?.top?r.top-36:0,o.fabAutoRepositioningTop<.5*window.innerHeight)return;n(o.fabAutoRepositioningTop),sendAnalyticsEvent([["DCBrowserExt:SidePanel:FabIcon:OverlappingRepositioned"]]),sendAnalyticsEvent([["DCBrowserExt:SidePanel:FabIcon:OverlappingRepositioned:ARF"]])}catch(t){chrome.runtime.sendMessage({main_op:"log-error",log:{message:"Error in repositioning the FAB if overlapping",error:t.toString()}})}},[e,o,n]),setFABTop:n}};