import React from "react"
import PropTypes from "prop-types"
import Helmet from "react-helmet"
import { useStaticQuery, graphql } from "gatsby"

function downloadLinks({ }) {
  const { site } = useStaticQuery(
    graphql`
    query {
      allFile {
        nodes {
          publicURL
        }
      }
    }
    `
  )
  return (
    <div>
      {site.allFile.nodes.publicURL}
    </div>
  );
}
export default downloadLinks();
