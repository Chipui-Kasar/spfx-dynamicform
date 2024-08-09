import * as React from "react";
import type { IDynamicFormProps } from "./IDynamicFormProps";
import { DynamicForm } from "@pnp/spfx-controls-react/lib/DynamicForm";
import styles from "./DynamicForm.module.scss";
import { Placeholder } from "@pnp/spfx-controls-react";

const DynamicForms: React.FC<IDynamicFormProps> = (
  props: IDynamicFormProps
) => {
  return (
    <section className={styles.dynamicForm}>
      {props.lists ? (
        <>
          <h2>{props.listName}</h2>
          <DynamicForm
            context={props.context}
            listId={props.lists}
            onCancelled={() => {
              console.log("Cancelled");
            }}
            onBeforeSubmit={async (listItem) => {
              return false;
            }}
            onSubmitError={(listItem, error) => {
              alert(error.message);
            }}
            onSubmitted={async (listItemData) => {
              console.log(listItemData);
            }}
          />
        </>
      ) : (
        <>
          <Placeholder
            iconName="Edit"
            iconText="Configure Page Hierarchy Web Part"
            description="Please configure the web part."
            buttonLabel="Configure"
            onConfigure={props.onConfigure}
            contentClassName={styles.placeHolder}
          />
        </>
      )}
    </section>
  );
};

export default DynamicForms;
