module.exports = {
  siteMetadata: {
    title: "Excel",
    description: "Excel Files, Notes, Tutorials etc.",
    author: {
      name: "Patrick Powell",
      email: "prp1277@gmail.com",
      handle: "prp1277",
      twitter: "http://bit.ly/Powell-Twitter",
      linkedin: "http://bit.ly/powell-linkedin",
      spotify: "https://spoti.fi/2L0Dt5m",
      github: "http://bit.ly/Powell-GitHub",
      reddit: "http://bit.ly/Powell-Reddit"
    },
  },
  plugins: [
    `gatsby-plugin-react-helmet`,
    {
      resolve: `gatsby-source-filesystem`,
      options: {
        name: `images`,
        path: `${__dirname}/src/images`,
      },
    },
    `gatsby-transformer-sharp`,
    `gatsby-plugin-sharp`,
    {
      resolve: `gatsby-plugin-manifest`,
      options: {
        name: `gatsby-starter-default`,
        short_name: `starter`,
        start_url: `/`,
        background_color: `#663399`,
        theme_color: `#663399`,
        display: `minimal-ui`,
        icon: `src/images/gatsby-icon.png`, // This path is relative to the root of the site.
      },
    },
    // this (optional) plugin enables Progressive Web App + Offline functionality
    // To learn more, visit: https://gatsby.dev/offline
    "gatsby-plugin-offline"
  ],
}
