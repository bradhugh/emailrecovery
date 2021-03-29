import * as React from "react";
import { Footer } from "./components/Footer";

interface IMailToProps
{
  address: string;
  text?: string;
}

const MailTo: React.FC<IMailToProps> = ({ address, text }) =>
{
  return <a href={"mailto:" + address}>{text ?? address}</a>
}

export const Privacy: React.FC = () => {
  return (
    <div>
      <h1>Privacy</h1>
      <div>
        The Email Recovery Outlook App does not collect or transmit any user
        information. Please refer any question to{" "}
        <MailTo address="emailrecovery@hughesonline.us" />
        .
      </div>
      <Footer />
    </div>
  );
};
