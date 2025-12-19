document.getElementById("fillA").addEventListener("click", async () => {
  fillLogin("Holden_Chen", "Cjz@10125011");
});

document.getElementById("fillB").addEventListener("click", async () => {
  fillLogin("dacsmy222", "222");
});

function fillLogin(username, password) {
  chrome.tabs.query({ active: true, currentWindow: true }, ([tab]) => {
    chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: (username, password) => {
        const usernameInput = document.querySelector('input#login');
        const passwordInput = document.querySelector('input#pin');

        if (usernameInput) {
          usernameInput.value = username;
          usernameInput.dispatchEvent(new Event("input", { bubbles: true }));
        }

        if (passwordInput) {
          passwordInput.value = password;
          passwordInput.dispatchEvent(new Event("input", { bubbles: true }));
        }
      },
      args: [username, password]
    });
  });
}
