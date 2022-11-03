import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
// import HeroList from "./HeroList";
import Progress from "./Progress";
import ItemsList from "./ItemsList";
const { Configuration, OpenAIApi } = require("openai");
/* global require, Word */


const configuration = new Configuration({
  apiKey: "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",
});

const openai = new OpenAIApi(configuration);

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      items: [],
      isClicked: "no",
      DataisLoaded: false,
    };
  }

  // componentDidMount() {
  // }

  click = async () => {
    /* providing token in bearer */
    // const response = await fetch("https://jsonplaceholder.typicode.com/users");
    // const jsonResponse = await response.json();
    return Word.run(async (context) => {
      const paragraphs = context.document.getSelection().paragraphs;
      paragraphs.load();
      await context.sync();
      var doc_text = paragraphs.items[0].text || "";
      var op_data = [];

      if (doc_text.length > 0) {
        const response_2 = await openai.createCompletion({
          model: "text-davinci-002",
          prompt: doc_text,
          temperature: 0,
          max_tokens: 6,
        });
        // const jsonResponse = await response_2.json();
        var { choices } = { ...response_2.data };
        op_data = choices.filter((d) => d.text != "");
      } else {
        op_data = [];
      }

      this.setState({
        items: op_data,
        // items: jsonResponse,
        isClicked: "yes",
      });
      // return React.createElement('h1', null, JSON.stringify(jsonResponse))
      this.render();
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
        <Header
          logo={require("./../../../assets/logo-filled.png")}
          title={this.props.title}
          message="NVS Text Suggestions"
        />
        <ItemsList suggestions={this.state.items} clicked={this.state.isClicked}>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Get suggestions
          </DefaultButton>
        </ItemsList>
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
