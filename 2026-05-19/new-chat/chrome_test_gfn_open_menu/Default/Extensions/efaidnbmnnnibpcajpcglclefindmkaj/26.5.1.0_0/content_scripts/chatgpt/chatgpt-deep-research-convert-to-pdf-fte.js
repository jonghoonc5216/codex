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
const CHATGPT_DEEP_RESEARCH_FTE_STORAGE_KEY="acrobat-chatgpt-convert-to-pdf-deep_research-fte";let chatgptDeepResearchTouchPointAddedResolve,chatgptDeepResearchTouchPointAddedPromise=new Promise(e=>{chatgptDeepResearchTouchPointAddedResolve=e});class ChatGPTDeepResearchConvertToPdfFte{id="chatgptDeepResearchConvertToPdfFte";timeout=2e3;static DOMAINS=["chatgpt.com"];constructor(){const e=window.location.hostname;if(!ChatGPTDeepResearchConvertToPdfFte.DOMAINS.some(t=>e.includes(t)))return this.isEligible=async()=>!1,void(this.render=async()=>{});this.initPromise=this.loadServices()}async loadServices(){this.fteUtils=await import(chrome.runtime.getURL("content_scripts/utils/fte-utils.js"))}async render(){this.pendingConfig&&this.frameId&&chrome.runtime.sendMessage({main_op:"chatgpt-deep-research-fte-render",frameId:this.frameId})}async isEligible(){const e=await chrome.runtime.sendMessage({main_op:"chatgpt-convert-to-pdf-init"}),t=e?.deepResearch;if(!t?.enableFte)return!1;await this.initPromise;const{frameId:i}=await chatgptDeepResearchTouchPointAddedPromise;this.frameId=i;const s=await this.fteUtils.initFteStateAndConfig(CHATGPT_DEEP_RESEARCH_FTE_STORAGE_KEY),a=await this.fteUtils.shouldShowFteTooltip(t.fteConfig,s,t.enableFte);return a&&(this.pendingConfig=t),a}}chrome.runtime.onMessage.addListener(e=>{"added-chatgpt-deep-research-touch-point"===e?.type&&chatgptDeepResearchTouchPointAddedResolve({frameId:e.frameId})}),window.ChatGPTDeepResearchConvertToPdfFte=ChatGPTDeepResearchConvertToPdfFte;