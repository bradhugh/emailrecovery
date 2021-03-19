import { ChoiceGroup, IChoiceGroupOption, PrimaryButton } from "@fluentui/react";
import React, { useEffect, useState } from "react";
import { useDispatch } from "react-redux";
import { ProgressComponent } from "../progress/ProgressComponent";
import { loadFolderHierarchyAsync, performRecoveryAsync } from "./recoverySlice";

const options: IChoiceGroupOption[] = [
  { key: 'recoverableitemsdeletions', text: 'Recoverable Items' },
  { key: 'recoverableitemspurges', text: 'Purges' },
];

export const RecoveryComponent = () =>
{
  const dispatch = useDispatch();
  const [selected, setSelected] = useState<string>("recoverableitemsdeletions");
  useEffect(() => void dispatch(loadFolderHierarchyAsync()), [dispatch]);

  return <>
      <ChoiceGroup options={options} required={true} selectedKey={selected} onChange={(e, opt) => opt?.key ? setSelected(opt.key) : null} />
      <PrimaryButton onClick={() => dispatch(performRecoveryAsync(selected))}>Start Recovery</PrimaryButton>
      <ProgressComponent />
      </>;
};