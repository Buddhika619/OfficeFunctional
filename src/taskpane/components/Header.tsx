import * as React from "react";
import { useRef } from "react";

// export interface HeaderProps {
//   title: string;
//   logo: string;
//   message: string;
// }

// export default class Header extends React.Component<HeaderProps> {
//   render() {
//     const { title, logo, message } = this.props;
//     const myRef = useRef<number | null>(null);
//     myRef.current = 3333;

//     return (
//       <section className="ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500">
//         <img width="90" height="90" src={logo} alt={title} title={title} />
//         <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{message}</h1>
//       </section>
//     );
//   }
// }

const Header = () => {
  const myRef = useRef<number | null>(null);
  myRef.current = 3333;
  return <div>{myRef.current}</div>;
};

export default Header;
