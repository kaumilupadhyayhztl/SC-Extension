// Open the side panel automatically when the extension icon is clicked
chrome.sidePanel
  .setPanelBehavior({ openPanelOnActionClick: true })
  .catch(console.error);

// Notify the side panel whenever a tab finishes loading.
// This lets the side panel re-inject the content script and keep
// "selection mode" alive after the user navigates to a new page.
chrome.tabs.onUpdated.addListener((tabId, changeInfo) => {
  if (changeInfo.status === 'complete') {
    chrome.runtime
      .sendMessage({ type: 'TAB_LOADED', tabId })
      .catch(() => {}); // side panel may not be open — ignore
  }
});
