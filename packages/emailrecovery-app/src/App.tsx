import React, { useEffect, useState } from "react";
import { RecoveryComponent } from "./features/recovery/RecoveryComponent";
import { Utils } from "./Utils";

const waitForMailboxAsync: () => Promise<void> = async () => {
  while (!Office?.context?.mailbox) {
    await Utils.setTimeoutAsync(100);
  }
};

function App() {
  const officeJsCdn =
    "https://appsforoffice.microsoft.com/lib/1/hosted/Office.js";

  const [isLoaded, setIsLoaded] = useState(false);

  useEffect(() => {
    if (!document.querySelector(`script[src="${officeJsCdn}"]`))
    {
      const elem = document.createElement("script");
      elem.src = officeJsCdn;
      elem.onload = () => {
        Office.initialize = () => {
          waitForMailboxAsync().then(() => setIsLoaded(true));
          console.log("Mailbox is: ", Office.context.mailbox);
        };
      };

      document.head.appendChild(elem);
    }
  }, []);

  if (!isLoaded) {
    return <div>Loading...</div>;
  }

  return (
    <div className="App">
      <RecoveryComponent />
    </div>
  );
}

export default App;
