import * as React from "react";
import { Button, ButtonType, TextField } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
/* global Button Header, HeroList, HeroListItem, Progress, Word */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      textValue: ""
    };
  }

  click = async () => {
    return Word.run(async context => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const selection = context.document.getSelection();
      const text = this.state.textValue;
      const regex = /(T-\d+-CR-\d+-\d+ (\d+[- ]?)+[a-zA-Z][a-zA-Z])|(ct[#]\d+ S(\.\d+)+(\(.\))* CC \d\d\/\d\d\/\d+)/g;
      const lineBreakRegex = /(T-\d+-CR-\d+-\d+ (\d+[- ]?)+[a-zA-Z][a-zA-Z])/g;
      const matches = text.match(regex);
      console.log("MATCHES");
      if (matches === null) {
        console.log("No matches")
        return;
      }

      var insertingText = "";
      matches.forEach((value, index) => {
        if (value.match(lineBreakRegex)) {
          if (index > 0) {
            insertingText += "\n"
          }
        }
        insertingText += value;
        insertingText += "\n";
      });

      const clipboard = selection.insertText(insertingText, Word.InsertLocation.end);
      clipboard.font.set({
        name: "Georgia",
        size: 12
      });

      this.setState({textValue: ""});

      await context.sync();
    });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        <HeroList message="Paste text below to format it" items={[]}>
          <TextField
            value={this.state.textValue}
            multiline
            rows={8}
            onChange={(e, val) => {
              this.setState({textValue: val})
            }}
          />
          <br />
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Paste to document
          </Button>
        </HeroList>
      </div>
    );
  }
}
