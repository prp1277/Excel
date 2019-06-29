---
title: Queries
date: 2019-06-29
---

## Directory Query

```gql
  query {
  allFile {
    edges {
      node {
        sourceInstanceName
        relativeDirectory
        relativePath
        publicURL
        id
      }
    }
  }
}
```
