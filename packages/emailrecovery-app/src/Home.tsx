import { Footer } from "./components/Footer";
import * as React from "react";

export const Home: React.FC = () => {
  return (
    <div>
      <h1>Email Recovery Outlook App</h1>
      <div>
        The Email Recovery Outlook App is provided to assist users with recovery
        of deleted items in a mailbox. The app is provided as-is with no
        warranty expressed or implied. That said, if you find bugs, please let
        us know at{" "}
        <a href="mailto:emailrecovery@hughesonline.us">
          emailrecovery@hughesonline.us
        </a>
        .
      </div>
      <Footer />
    </div>
  );
}
