import { apiRunner, apiRunnerAsync } from "./api-runner-browser"
import React, { createElement } from "react"
import ReactDOM from "react-dom"
import { Router, navigate } from "@reach/router"
import { ScrollContext } from "gatsby-react-router-scroll"
import domReady from "@mikaelkristiansson/domready"
import {
  shouldUpdateScroll,
  init as navigationInit,
  RouteUpdates,
} from "./navigation"
import emitter from "./emitter"
import PageRenderer from "./page-renderer"
import asyncRequires from "./async-requires"
import matchPaths from "./match-paths.json"
import loader, { setApiRunnerForLoader } from "./loader"
import EnsureResources from "./ensure-resources"
import stripPrefix from "./strip-prefix"

window.asyncRequires = asyncRequires
window.___emitter = emitter
window.___loader = loader
window.___webpackCompilationHash = window.webpackCompilationHash

loader.addProdRequires(asyncRequires)
loader.addMatchPaths(matchPaths)
setApiRunnerForLoader(apiRunner)

navigationInit()

// Let the site/plugins run code very early.
apiRunnerAsync(`onClientEntry`).then(() => {
  // Let plugins register a service worker. The plugin just needs
  // to return true.
  if (apiRunner(`registerServiceWorker`).length > 0) {
    require(`./register-service-worker`)
  }

  class RouteHandler extends React.Component {
    render() {
      let { location } = this.props

      return (
        <EnsureResources location={location}>
          {({ pageResources, location }) => (
            <RouteUpdates location={location}>
              <ScrollContext
                location={location}
                shouldUpdateScroll={shouldUpdateScroll}
              >
                <PageRenderer
                  {...this.props}
                  location={location}
                  pageResources={pageResources}
                  {...pageResources.json}
                />
              </ScrollContext>
            </RouteUpdates>
          )}
        </EnsureResources>
      )
    }
  }

  const { pagePath, location: browserLoc } = window

  // Explicitly call navigate if the canonical path (window.pagePath)
  // is different to the browser path (window.location.pathname). But
  // only if NONE of the following conditions hold:
  //
  // - The url matches a client side route (page.matchPath)
  // - it's a 404 page
  // - it's the offline plugin shell (/offline-plugin-app-shell-fallback/)
  if (
    pagePath &&
    __BASE_PATH__ + pagePath !== browserLoc.pathname &&
    !(
      loader.findMatchPath(stripPrefix(browserLoc.pathname, __BASE_PATH__)) ||
      pagePath === `/404.html` ||
      pagePath.match(/^\/404\/?$/) ||
      pagePath.match(/^\/offline-plugin-app-shell-fallback\/?$/)
    )
  ) {
    navigate(__BASE_PATH__ + pagePath + browserLoc.search + browserLoc.hash, {
      replace: true,
    })
  }

  loader.loadPage(browserLoc.pathname).then(() => {
    const Root = () =>
      createElement(
        Router,
        {
          basepath: __BASE_PATH__,
        },
        createElement(RouteHandler, { path: `/*` })
      )

    const WrappedRoot = apiRunner(
      `wrapRootElement`,
      { element: <Root /> },
      <Root />,
      ({ result }) => {
        return { element: result }
      }
    ).pop()

    let NewRoot = () => WrappedRoot

    const renderer = apiRunner(
      `replaceHydrateFunction`,
      undefined,
      ReactDOM.hydrate
    )[0]

    domReady(() => {
      renderer(
        <NewRoot />,
        typeof window !== `undefined`
          ? document.getElementById(`___gatsby`)
          : void 0,
        () => {
          apiRunner(`onInitialClientRender`)
        }
      )
    })
  })
})
