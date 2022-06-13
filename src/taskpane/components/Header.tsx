import * as React from "react";

export interface HeaderProps {
  title: string;
  logo: string;
  message: string;
}

const Header: React.FC<HeaderProps> = ({ title, logo, message }) => {
  const render = () => {

    return (
      <section className="ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500">
        <img width="90" height="90" src={logo} alt={title} title={title} />
        <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{message}</h1>
      </section>
    );
  }

  return render();
}

export default Header;