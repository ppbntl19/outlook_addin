!function(e){var t={};function n(r){if(t[r])return t[r].exports;var c=t[r]={i:r,l:!1,exports:{}};return e[r].call(c.exports,c,c.exports,n),c.l=!0,c.exports}n.m=e,n.c=t,n.d=function(e,t,r){n.o(e,t)||Object.defineProperty(e,t,{enumerable:!0,get:r})},n.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},n.t=function(e,t){if(1&t&&(e=n(e)),8&t)return e;if(4&t&&"object"==typeof e&&e&&e.__esModule)return e;var r=Object.create(null);if(n.r(r),Object.defineProperty(r,"default",{enumerable:!0,value:e}),2&t&&"string"!=typeof e)for(var c in e)n.d(r,c,function(t){return e[t]}.bind(null,c));return r},n.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(t,"a",t),t},n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},n.p="",n(n.s=307)}({307:function(e,t,n){(function(e){function t(e,t,n,r,c,o,i){try{var s=e[o](i),a=s.value}catch(e){return void n(e)}s.done?t(a):Promise.resolve(a).then(r,c)}function n(e){return function(){var n=this,r=arguments;return new Promise((function(c,o){var i=e.apply(n,r);function s(e){t(i,c,o,s,a,"next",e)}function a(e){t(i,c,o,s,a,"throw",e)}s(void 0)}))}}var r={};function c(){return(c=n(regeneratorRuntime.mark((function e(t){var n;return regeneratorRuntime.wrap((function(e){for(;;)switch(e.prev=e.next){case 0:if(r.default_classfication=Office.context.roamingSettings.get("default_classfication"),n=Office.context.roamingSettings.get("quick_classfication"),!r.default_classfication){e.next=14;break}return e.prev=3,e.next=6,u(r.default_classfication);case 6:setTimeout((function(){t.completed({allowEvent:!0})}),2e3),e.next=12;break;case 9:e.prev=9,e.t0=e.catch(3),l(e.t0,"error","validateClassfication");case 12:e.next=15;break;case 14:n?(p(),setTimeout((function(){t.completed({allowEvent:!0})}),2e3)):(l(" Please set a classification for this email.[validateClassfication]","error","validateClassfication"),t.completed({allowEvent:!1}));case 15:case"end":return e.stop()}}),e,null,[[3,9]])})))).apply(this,arguments)}function o(){return(o=n(regeneratorRuntime.mark((function e(t){return regeneratorRuntime.wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,u("Internal");case 3:d(),t.completed(),e.next=12;break;case 7:e.prev=7,e.t0=e.catch(0),p(),l(e.t0,"error","setClassfication"),t.completed();case 12:case"end":return e.stop()}}),e,null,[[0,7]])})))).apply(this,arguments)}function i(){return(i=n(regeneratorRuntime.mark((function e(t){return regeneratorRuntime.wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,u("Screte");case 3:d(),t.completed(),e.next=12;break;case 7:e.prev=7,e.t0=e.catch(0),p(),l(e.t0,"error","setClassfication"),t.completed();case 12:case"end":return e.stop()}}),e,null,[[0,7]])})))).apply(this,arguments)}function s(){return(s=n(regeneratorRuntime.mark((function e(t){return regeneratorRuntime.wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,u("Confidential");case 3:d(),t.completed(),e.next=12;break;case 7:e.prev=7,e.t0=e.catch(0),p(),l(e.t0,"error","setClassfication"),t.completed();case 12:case"end":return e.stop()}}),e,null,[[0,7]])})))).apply(this,arguments)}function a(){return(a=n(regeneratorRuntime.mark((function e(t){return regeneratorRuntime.wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,u("Public");case 3:d(),t.completed(),e.next=12;break;case 7:e.prev=7,e.t0=e.catch(0),t.completed(),p(),l(e.t0,"error","setClassfication");case 12:case"end":return e.stop()}}),e,null,[[0,7]])})))).apply(this,arguments)}function u(e){return f.apply(this,arguments)}function f(){return(f=n(regeneratorRuntime.mark((function e(t){var n,r,c,o,i,s;return regeneratorRuntime.wrap((function(e){for(;;)switch(e.prev=e.next){case 0:n=Office.context.mailbox.userProfile.emailAddress,r=Office.context.mailbox.userProfile.timeZone,c="This email is classified as "+t+". By "+n+" at "+(new Date).toLocaleDateString()+" "+r,o="This email is classified ["+t+"]",i="This email is classified ["+t+"]",e.prev=5,s=[{displayName:t,color:Office.MailboxEnums.CategoryColor.Preset0}],Office.context.mailbox.masterCategories.addAsync(s,(function(e){e.status===Office.AsyncResultStatus.Succeeded?Office.context.mailbox.item.categories.addAsync([t],(function(e){e.status===Office.AsyncResultStatus.Succeeded?l("Successfully added categories","success","category"):l("categories.addAsync call failed with error: "+e.error.message,"error","category")})):(Office.context.mailbox.item.categories.addAsync([t],(function(e){e.status===Office.AsyncResultStatus.Succeeded?l("Successfully added categories","success","category"):l("categories.addAsync call failed with error: "+e.error.message,"error","category")})),l("Unable to set the category MasterCategories.addAsync"+e.error.message,"error","Createcategory"))})),Office.context.mailbox.item.subject.setAsync(o,(function(e){e.status===Office.AsyncResultStatus.Succeeded?l("Successfully added subject","success","subject"):l("Unable to set the subject: "+e.error.message,"error","subject")})),Office.context.mailbox.item.body.prependAsync(c,{coercionType:"html"},(function(e){e.status===Office.AsyncResultStatus.Succeeded?(l("Successfully added body","success","body"),Office.context.mailbox.item.body.getAsync("html",(function(e){if(e.status===Office.AsyncResultStatus.Succeeded){var t=e.value+"</br></br></br><h4>"+i+"</h4>";Office.context.mailbox.item.body.setAsync(t,{coercionType:"html"},(function(e){e.status===Office.AsyncResultStatus.Succeeded?l("Successfully added footer","success","footer"):l("body.setAsync call failed with error: "+e.error.message,"error","footer")}))}else l("Unable to set the footer: "+e.error.message,"error","footer")}))):l("Unable to set the body: "+e.error.message,"error","body")})),e.next=16;break;case 12:throw e.prev=12,e.t0=e.catch(5),Office.context.mailbox.item.body.setAsync("<b>"+e.t0+"</b>",{coercionType:"html"},(function(e){})),TypeError("SetClassification Error");case 16:case"end":return e.stop()}}),e,null,[[5,12]])})))).apply(this,arguments)}function l(e,t,n){t&&"error"==t?Office.context.mailbox.item.notificationMessages.addAsync(n,{type:"errorMessage",message:e}):Office.context.mailbox.item.notificationMessages.addAsync(n,{type:"informationalMessage",message:e,icon:"iconid",persistent:!1})}function d(){Office.context.roamingSettings.set("quick_classfication",!0),Office.context.roamingSettings.saveAsync()}function p(e){Office.context.roamingSettings.remove("quick_classfication"),Office.context.roamingSettings.saveAsync()}Office.initialize=function(e){mailboxItem=Office.context.mailbox.item};var m="undefined"!=typeof self?self:"undefined"!=typeof window?window:void 0!==e?e:void 0;m.setInternal=function(e){return o.apply(this,arguments)},m.setConfidential=function(e){return s.apply(this,arguments)},m.setPublic=function(e){return a.apply(this,arguments)},m.setScrete=function(e){return i.apply(this,arguments)},m.validateClassfication=function(e){return c.apply(this,arguments)}}).call(this,n(308))},308:function(e,t){var n;n=function(){return this}();try{n=n||new Function("return this")()}catch(e){"object"==typeof window&&(n=window)}e.exports=n}});
//# sourceMappingURL=commands.js.map