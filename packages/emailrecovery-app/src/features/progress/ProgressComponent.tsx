import React from "react";
import { useId } from "@uifabric/react-hooks";
import {
  getTheme,
  mergeStyleSets,
  FontWeights,
  Modal,
  ProgressIndicator,
} from "@fluentui/react";
import { useSelector } from "react-redux";
import { RootState } from "../../app/store";
import { Utils } from "../../Utils";

const theme = getTheme();
const contentStyles = mergeStyleSets({
  container: {
    display: "flex",
    flexFlow: "column nowrap",
    alignItems: "stretch",
  },
  header: [
    theme.fonts.xLargePlus,
    {
      flex: "1 1 auto",
      borderTop: `4px solid ${theme.palette.themePrimary}`,
      color: theme.palette.neutralPrimary,
      display: "flex",
      alignItems: "center",
      fontWeight: FontWeights.semibold,
      padding: "12px 12px 14px 24px",
    },
  ],
  body: {
    flex: "4 4 auto",
    padding: "0 24px 24px 24px",
    overflowY: "hidden",
    selectors: {
      p: { margin: "14px 0" },
      "p:first-child": { marginTop: 0 },
      "p:last-child": { marginBottom: 0 },
    },
  },
});

export const ProgressComponent: React.FC = () => {
  const { active, activity, status } = useSelector(
    (state: RootState) => state.progress
  );

  // Use useId() to ensure that the IDs are unique on the page.
  // (It's also okay to use plain strings and manually ensure uniqueness.)
  const titleId = useId("title");

  return (
    <Modal
      titleAriaId={titleId}
      isOpen={active}
      isBlocking={true}
      containerClassName={contentStyles.container}
    >
      <div className={contentStyles.header}>
        <span id={titleId}>Please wait.</span>
      </div>
      <div className={contentStyles.body}>
        <div
          dangerouslySetInnerHTML={{
            __html: Utils.sanitizeHtmlDefault(activity),
          }}
        ></div>
        <ProgressIndicator description={status} />
      </div>
    </Modal>
  );
};
