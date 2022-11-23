import {
  teamsDarkTheme,
  teamsHighContrastTheme,
  teamsLightTheme
} from "@fluentui/react-components";

enum Theme {
  Light = "default",
  Dark = "dark",
  HighContrast = "contrast",
}

export const getThemeFromThemeString = (t: string) => {
  switch (t) {
    case Theme.Dark:
      return teamsDarkTheme;
    case Theme.HighContrast:
      return teamsHighContrastTheme;
    default:
      return teamsLightTheme;
  }
};
