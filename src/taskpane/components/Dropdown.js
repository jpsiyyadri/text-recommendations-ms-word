import React from "react";
import PropTypes from "prop-types";

export default function Dropdown({ options, selected, onClick }) {
  return (
    <select onChange={(e) => onClick(e.target.value)} defaultValue={selected}>
      {options.map((option, idx) => {
        // if(option === selected){
        //     return option === selected?<option key={idx} selected value={option}>{option}</option>:<option key={idx} value={option}>{option}</option>
        // }
        return <option key={idx}>{option}</option>;
      })}
    </select>
  );
}

Dropdown.propTypes = {
  options: PropTypes.array,
  selected: PropTypes.string,
  onClick: PropTypes.func,
};
