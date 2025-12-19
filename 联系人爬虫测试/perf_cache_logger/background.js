chrome.webRequest.onCompleted.addListener(
  function (details) {
    const url = new URL(details.url);
    const isJSON = details.responseHeaders?.some(h =>
      h.name.toLowerCase() === 'content-type' &&
      h.value?.includes('application/json')
    );

    if (isJSON || url.pathname.endsWith('.json')) {
      const pathSegments = url.pathname.split('/');
      const logicalName = pathSegments[pathSegments.length - 1] || '[未知文件]';
      console.log(`[XHR JSON] 请求逻辑文件名: ${logicalName}`);
    }
  },
  { urls: ["<all_urls>"], types: ["xmlhttprequest"] },
  ["responseHeaders"]
);
