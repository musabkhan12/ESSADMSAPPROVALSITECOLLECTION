declare module "*.png";
declare module "*.jpg";
declare module "*.jpeg";
declare module "*.svg";
declare module "*.gif";

// ✅ Add these for CSS and SCSS
declare module "*.css" {
  const content: { [className: string]: string };
  export default content;
}

declare module "*.scss" {
  const content: { [className: string]: string };
  export default content;
}

// ✅ Add these for font files
declare module "*.ttf";
declare module "*.woff";
declare module "*.woff2";