import { Separator } from "@fluentui/react";
import * as React from "react";
import { Link } from "react-router-dom";

export const Footer: React.FC = () => {
  return (
    <>
      <Separator />
      <div>© 2021 - Brad Hughes{ " | " }<Link to="/privacy">Privacy</Link></div>
    </>);
};
