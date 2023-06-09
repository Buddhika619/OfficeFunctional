import React, { ReactElement } from "react";

/* global console,fetch, Office, */

export interface HeroListItem {
  icon: string;
  primaryText: string;
  url: string;
}

export interface HeroListProps {
  message: string;
  items: HeroListItem[];
  children: React.ReactNode;
}

const HeroList = ({ items, children }: HeroListProps): ReactElement => {
  const getBase64FromUrl = async (url: string): Promise<string> => {
    const data = await fetch(url);
    const blob = await data.blob();
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.readAsDataURL(blob);
      reader.onloadend = () => {
        const base64data = reader.result as string;
        resolve(base64data);
      };
    });
  };
  const addImageToSlide = async (url: string) => {
    // let encodedImage =
    //   "https://cdn.glitch.global/04ac2eab-7093-47ad-976f-739938dcbb74/pulp-fiction-john-travolta.gif?v=1667742225377";

    // Create a new Image object
    const res = (await getBase64FromUrl(url)) as string;
    let image = "";
    if (res) {
      image = res.slice(22, res.length - 1);
    }
    // Set the source of the image to the base64-encoded string

    console.log("image");
    Office.context.document.setSelectedDataAsync(image, {
      coercionType: Office.CoercionType.Image,
      imageLeft: 50,
      imageTop: 50,
      imageWidth: 250,
      imageHeight: 250,
    });
  };

  const listItems = items.map((item, index) => (
    <li className="ms-ListItem" key={index}>
      <img src={item.url} alt="" width="150px" onClick={() => addImageToSlide(item.url)} />
    </li>
  ));

  return (
    <main className="ms-welcome__main">
      {/* <h2 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">{message}</h2> */}
      <ul className="ms-List ms-welcome__features ms-u-slideUpIn10">{listItems}</ul>
      <img src="" alt="" />
      {children}
    </main>
  );
};

export default HeroList;
