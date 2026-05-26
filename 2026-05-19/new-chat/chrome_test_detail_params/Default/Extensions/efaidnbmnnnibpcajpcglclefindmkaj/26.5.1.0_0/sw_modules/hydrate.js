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
import{floodgate as t}from"./floodgate.js";import{communicate as s}from"./communicate.js";import{dcLocalStorage as r,callWithStorage as n}from"../common/local-storage.js";import{common as e}from"./common.js";import{migrateLegacyKeysToStable as o,FLOODGATE_KEY as a,COMMUNICATE_KEY as i,COMMON_KEY as c,USE_STABLE_FLAG as m,HYDRATE_TS_KEY as h,HYDRATE_STABLE_MIGRATION_LS as f}from"./migrate-legacy-keys.js";function u(){return"true"===r.getItem(m)||!0===r.getItem(f)}function y(t,s){return u()&&r.getItem(m)?s:t.constructor.name}const l=new function(){this.instances=[[t,a],[s,i],[e,c]],this.status={},this.instances.forEach(([,t])=>{this.status[t]=!1}),this.syncInterval=null,this.do=()=>{u()&&o(),this.instances.forEach(([t,s])=>{if(!this.status[s]){const n=y(t,s),e=r.getItem(n)||"[]";let o=[];try{const t=JSON.parse(e);Array.isArray(t)?o=t:Array.isArray(t?.data)&&(o=t.data)}catch(t){o=[]}o.forEach(([s,r])=>{t[s]=r}),this.status[s]=!0}}),this.sync()},this.sync=()=>{this.syncInterval||(this.syncInterval=setInterval(()=>{this.instances.forEach(([t,s])=>{t[h]=Date.now();const n=y(t,s);r.setItem(n,JSON.stringify(Object.entries(t)))})},1e3))}},g=(...t)=>n(l.do,...t);export{g as hydrateWithStorage};