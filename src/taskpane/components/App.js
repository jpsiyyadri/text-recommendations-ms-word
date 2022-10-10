import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
// import HeroList from "./HeroList";
import Progress from "./Progress";
import ItemsList from "./ItemsList";

/* global require, fetch */

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
    const response = await fetch("https://jsonplaceholder.typicode.com/users");
    const jsonResponse = await response.json();

    this.setState({
      items: jsonResponse,
      isClicked: "yes",
    });
    // return React.createElement('h1', null, JSON.stringify(jsonResponse))
    this.render();
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
