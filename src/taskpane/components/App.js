import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";
const { Configuration, OpenAIApi } = require("openai");
import Suggestions from "./Suggestions";
import Dropdown from "./Dropdown";

/* global Word, require */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    const configuration = new Configuration({
      apiKey: "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",
    });
    this.state = {
      listItems: [],
      model_names: {
        "text-davinci-002": "text-davinci-002",
        "text-curie-001": "text-curie-001",
      },
      model: "text-davinci-002",
      isClicked: "no",
    };
    this.openai = new OpenAIApi(configuration);
  }

  update_dropdown = (e) => {
    this.setState({
      model: e,
    });
  };

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
  }

  writeText = (e) => {
    const clicked_text = " " + e.target.dataset.val;
    return Word.run(async (context) => {
      const paragraphs = context.document.getSelection().paragraphs;
      paragraphs.load();
      await context.sync();
      paragraphs.items[0].insertText(clicked_text, Word.InsertLocation.end);
      this.setState({
        listItems: [],
        isClicked: "yes",
      });
      this.render();
    });
  };

  click = async () => {
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
      const paragraphs = context.document.getSelection().paragraphs;
      paragraphs.load();
      await context.sync();
      var doc_text = paragraphs.items[0].text || "";
      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      const response_2 = await this.openai.createCompletion({
        model: this.state.model_names[this.state.model],
        prompt: doc_text,
        temperature: 0,
        max_tokens: 6,
        n: 3,
      });
      let { choices } = { ...response_2.data };
      this.setState({
        listItems: choices,
      });

      await context.sync();
    });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Run
          </DefaultButton>
        </HeroList>
        <Dropdown
          options={this.state.model_names}
          selected={this.state.model}
          onClick={this.update_dropdown}
        ></Dropdown>
        <Suggestions items={this.state.listItems}></Suggestions>
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
