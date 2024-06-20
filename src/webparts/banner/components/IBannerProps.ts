export interface IBannerProps {
  description: string;
  secondsSlick: number;
  checkbox: boolean;
  color: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userDisplayEmail: string;
  site: {
    absoluteUrl: string;
    serverRelativeUrl: string;
    serverRequestPath: string;
  };
}
