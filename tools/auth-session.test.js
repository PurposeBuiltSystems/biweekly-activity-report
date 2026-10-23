'use strict';
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
function load(pca) {
 const context = {self:{}, msal:{createNestablePublicClientApplication:async()=>pca},setTimeout,clearTimeout,Office:{context:{mailbox:{userProfile:{emailAddress:'sample@example.org'}},roamingSettings:{get(){return false},set(){},saveAsync(fn){fn()}}}}};
 context.window = context.self;
 vm.runInNewContext(fs.readFileSync(require('node:path').join(__dirname,'../src/graph.js'),'utf8').replace('  var GraphData = {', '  var GraphData = {getToken: getToken,'),context);
 return context.self.GraphData;
}
(async()=>{
 let popups=0;
 const g=load({getAllAccounts:()=>[],acquireTokenSilent:async()=>({accessToken:'token',account:{username:'sample@example.org'}}),acquireTokenPopup:async()=>{popups++;return {accessToken:'interactive'}}});
 assert.equal(await g.getToken(),'token');assert.equal(await g.currentAccount(),'sample@example.org');
 await g.signOut();assert.equal(await g.currentAccount(),null);assert.equal(await g.getToken(),'interactive');assert.equal(popups,1);
 let finish;let entered; const started=new Promise(r=>entered=r);
 const race=load({getAllAccounts:()=>[],acquireTokenSilent:()=>{entered();return new Promise(r=>finish=r)},acquireTokenPopup:async()=>{throw Error('must not open popup after sign out')}});
 const pending=race.getToken();await started;await race.signOut();finish({accessToken:'stale'});await assert.rejects(pending,/cancelled/);assert.equal(await race.currentAccount(),null);
 let completePopup; let popupEntered;
 const popupStarted = new Promise(r=>popupEntered=r);
 const popupRace=load({getAllAccounts:()=>[],acquireTokenSilent:async()=>{throw Error('interaction required')},acquireTokenPopup:()=>{popupEntered();return new Promise(r=>completePopup=r)}});
 const popupPending=popupRace.getToken(); await popupStarted; await popupRace.signOut(); completePopup({accessToken:'stale popup'});
 await assert.rejects(popupPending,/cancelled/); assert.equal(popupRace.isSignedOut(),true);
 console.log('Authentication state and sign-out race regressions passed.');
})().catch(e=>{console.error(e);process.exitCode=1});
