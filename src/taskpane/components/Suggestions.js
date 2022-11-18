import React from "react";
import PropTypes from "prop-types";

export const Suggestions = ({ items, onClick }) => {
  return (
    <div className="suggestions">
      {items.map((item, idx) => {
        return (
          <div className="suggestion" key={idx} data-idx={idx} onClick={(e) => onClick(e.target)}>
            {item.text}
          </div>
        );
      })}
    </div>
  );
};

Suggestions.defaultProps = {
  items: [1, 2, 3, 4],
};

Suggestions.propTypes = {
  items: PropTypes.array,
  onClick: PropTypes.func,
};

export default Suggestions;
