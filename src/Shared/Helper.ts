export function getUrlParameterValue(paramName:string) {
  const urlParams = new URLSearchParams(window.location.search);
  return urlParams.get(paramName);
}