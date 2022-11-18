import * as React from "react";
import PropTypes from "prop-types";
import Progress from "./Progress";
const { Configuration, OpenAIApi } = require("openai");
import Suggestions from "./Suggestions";
import Dropdown from "./Dropdown";
import { useState, useEffect } from "react";

/* global Word, require */

export default function App({ title, isOfficeInitialized }) {
  const configuration = new Configuration({
    // eslint-disable-next-line no-undef
    apiKey: "",
  });
  const model_names = {
    "text-davinci-002": "text-davinci-002",
    "text-curie-001": "text-curie-001",
  };
  const openai = new OpenAIApi(configuration);
  const [choices, setChoices] = useState([{ text: "hi" }, { text: "by" }, { text: "cu" }]);
  const [model, setModel] = useState("text-davinci-002");

  useEffect(() => {
    // Word.addEventListener("keydown", detectKeydown, true);
    // Word.Document.addEventListener("keydown", detectKeydown, true);
    // Word.Body.addEventListener("keydown", detectKeydown, true);
    return Word.run(async (context) => {
      context.addEventListener("keydown", detectKeydown, true);
    });
  });

  const detectKeydown = (e) => {
    setChoices([]);
    if (e.key == "Tab") {
      click();
    }
  };

  const click = async () => {
    return Word.run(async (context) => {
      const paragraphs = context.document.getSelection().paragraphs;
      paragraphs.load();
      await context.sync();
      var doc_text = paragraphs.items[0].text || "";
      const response_2 = await openai.createCompletion({
        model: model_names[model],
        prompt: doc_text,
        temperature: 0,
        max_tokens: 6,
        n: 3,
      });
      let { choices } = { ...response_2.data };
      choices = choices.filter((d) => d.text != "");
      choices.forEach((item) => {
        item.text = item.text.replaceAll("\n", " ");
      })
      setChoices(choices);
      await context.sync();
    });
  };

  const update_dropdown = (e) => {
    setModel(e);
    setChoices([]);
  };

  const write = (e) => {
    const idx = e.dataset.idx;
    const clicked_text = " " + choices[idx].text;
    return Word.run(async (context) => {
      const paragraphs = context.document.getSelection().paragraphs;
      paragraphs.load();
      await context.sync();
      paragraphs.items[0].insertText(clicked_text, Word.InsertLocation.end);
      setChoices([]);
      await context.sync();
    });
  };

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
      <div>
        <Dropdown
          options={["text-davinci-002", "text-curie-001"]}
          selected={model}
          onClick={update_dropdown}
          className="dropdown"
        ></Dropdown>
        {/* <Select options={["text-davinci-002", "text-curie-001"]} onChange={(values) => update_dropdown(values)} /> */}
      </div>
      <div className="align-items-center">
        <button className="btn" onClick={click}>Click</button>
      </div>
      <div>
        <Suggestions items={choices} onClick={write}></Suggestions>
      </div>
    </div>
  );
  //   }
}

App.propTypes = {
  isOfficeInitialized: PropTypes.bool,
  title: PropTypes.string,
};