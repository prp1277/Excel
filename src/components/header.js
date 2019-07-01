import { Link } from "gatsby"
import React from "react"

const Header = ({ siteTitle }) => (
  <header>
    <div className=""></div>
    <h1>
      <Link to="/">{siteTitle}</Link>
    </h1>
  </header>
)

export default Header
