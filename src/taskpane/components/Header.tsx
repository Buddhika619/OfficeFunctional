import React, { ReactElement, useRef } from "react";

export interface HeaderProps {
  title: string;
  logo: string;
  message: string;
}

const Header = (): ReactElement => {
  const myRef = useRef<String | null>(null);
  myRef.current = "Hello";

  return (
    <section className="ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500">
      <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{myRef}</h1>
    </section>
  );
};

export default Header;
