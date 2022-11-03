/* eslint-disable react/jsx-key */
import * as React from "react";
import PropTypes from "prop-types";
/* global Word */

export default class ItemsList extends React.Component {
  render() {
    const { children, suggestions, clicked } = this.props;

    // const listItems = suggestions.map(function (suggestion) {
    //   var suggestion_text = suggestion.text.replaceAll("\n", " ");
    //   return (
    //     <div data-val={suggestion_text} onClick={this.click} style={{ cursor: "pointer" }} className="suggestion-item">
    //       {suggestion_text}
    //     </div>
    //   );
    // });

    const listItems = suggestions.map((suggestion) => (
      <div
        data-val={suggestion.text.replaceAll("\n", " ")}
        onClick={this.click}
        style={{ cursor: "pointer" }}
        className="suggestion-item"
      >
        {suggestion.text.replaceAll("\n", " ")}
      </div>
    ));
    if (clicked == "no") {
      return <>{children}</>;
    }
    if (suggestions.length > 0) {
      return (
        <>
          Suggestions...
          <div className="suggestions-list">{listItems}</div>
          {children}
        </>
      );
    }
    return (
      <>
        <h1> No suggestions available </h1>
        {children}
      </>
    );
  }

  click = (e) => {
    const clicked_text = " " + e.target.dataset.val;
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */
      // insert a paragraph at the end of the document.
      // const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

      // change the paragraph color to blue.
      // paragraph.font.color = "blue";

      // await context.sync();

      const paragraphs = context.document.getSelection().paragraphs;
      paragraphs.load();
      await context.sync();
      paragraphs.items[0].insertText(clicked_text, Word.InsertLocation.end);
      // return <ItemsList suggestions={["abc", "def"]} />;
      // paragraphs.items[0].insertText(" New sentence in the paragraph.", Word.InsertLocation.end);
      // await context.sync();
      this.setState({
        items: [],
        isClicked: "yes",
      });
      this.render();
    });
  };
}

ItemsList.propTypes = {
  suggestions: PropTypes.array,
  children: PropTypes.node,
  clicked: PropTypes.string,
};
