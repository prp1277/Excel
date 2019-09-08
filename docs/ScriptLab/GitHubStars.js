/**
 * Gets the star count for a given org/user and repo. Try =GETSTARCOUNT("officedev","office-js")
 * @customfunction
 * @param userName Name of org or user.
 * @param repoName Name of the repo.
 * @return Number of stars.
 */
async function getStarCount(userName = "prp1277", repoName = "Excel") {
  //You can change this URL to any web request you want to work with.
  const url = `https://api.github.com/repos/${userName}/${repoName}`;
  const response = await fetch(url);

  //Expect that status code is in 200-299 range
  if (!response.ok) {
    throw new Error(response.statusText);
  }
  const jsonResponse = await response.json();
  return jsonResponse.watchers_count;
}
