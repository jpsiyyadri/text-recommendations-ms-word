import * as React from "react";
import PropTypes from "prop-types";

export default class Header extends React.Component {
  render() {
    const { title, logo, message } = this.props;

    return (
      <section className="ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500">
        <img width="20" height="20" src={logo} alt={title} title={title} />
        <h6 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{message}</h6>
      </section>
    );
  }
}

Header.propTypes = {
  title: PropTypes.string,
  logo: PropTypes.string,
  message: PropTypes.string,
};
