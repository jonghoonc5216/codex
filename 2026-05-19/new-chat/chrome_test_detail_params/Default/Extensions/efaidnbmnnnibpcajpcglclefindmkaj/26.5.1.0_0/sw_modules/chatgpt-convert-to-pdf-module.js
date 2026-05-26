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
import t from"./CacheStore.js";import{floodgate as o}from"./floodgate.js";import{checkUserLocaleEnabled as e,safeParseFeatureFlag as c}from"./gsuite/util.js";import{setExperimentCodeForAnalytics as n,removeExperimentCodeForAnalytics as r}from"../common/experimentUtils.js";import{common as a}from"./common.js";import{forceResetService as s}from"./force-reset-service.js";async function i(t){const[a,s]=await Promise.all([o.hasFlag("dc-cv-chatgpt-dr-convert-to-pdf-touchpoint"),o.hasFlag("dc-cv-chatgpt-dr-convert-to-pdf-touchpoint-control")]);let i;a?i=c("dc-cv-chatgpt-dr-convert-to-pdf-touchpoint"):s&&(i=c("dc-cv-chatgpt-dr-convert-to-pdf-touchpoint-control"));const l=e(i?.enLocaleEnabled,i?.nonEnLocaleEnabled);a&&l?(r("CDRC"),n("CDR")):s&&l?(r("CDR"),n("CDRC")):(r("CDR"),r("CDRC"));const p=i?.selectors||null,f=i?.tooltip||{};t({deepResearch:{enableChatGPTDeepResearchTouchpoint:a&&l,enableFte:i?.fteEnabled||!1,fteConfig:f,selectors:p}})}const l="chatgptHtmlToPdfConfig",p="chatgpt-html-to-pdf-config";async function f(){const o=a.getChatgptHtmlToPdfConfigUri();if(!o)return null;const e=new t(p);try{const t=async()=>{const t=await async function(t){const o=await fetch(t);if(!o.ok)return null;const e=await o.json();return e?.["deep-research"]??e??null}(o);return t&&await e.set(l,t),t},{executionResult:c}=await s.executeFeature(p,t);if(c)return c}catch(t){}return e.get(l)}export{i as chatgptConvertToPdfInit,f as getChatgptDeepResearchHtmlToPdfConfig};