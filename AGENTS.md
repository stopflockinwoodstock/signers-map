# Repository Rules

- Do not commit person-level signer data.
- Do not commit signer names, addresses, postal codes, emails, phone numbers, signature dates, or raw petition exports.
- Commit only aggregate signer outputs, such as city/state counts and city-level coordinates.
- Keep source petition workbooks and exports local unless they have been explicitly anonymized.
- Use `data/city_overrides.csv` only for aggregate city-level coordinate overrides; do not add signer rows or person-level data there.

## GitHub And Commit Identity

- All commits in this repository must use:
  - `user.name`: `Woodstock-IL-Cameras`
  - `user.email`: `317771352+Woodstock-IL-Cameras@users.noreply.github.com`
- Before committing, verify:
  - `git config --get user.name`
  - `git config --get user.email`
- GitHub CLI operations must use the `Woodstock-IL-Cameras` account. Before GitHub operations, verify `gh auth status` shows `Woodstock-IL-Cameras` as the active account.
- Do not commit, push, open PRs, or alter remotes using any other GitHub account or commit identity.
- Push only `main` to `origin` unless the user explicitly asks for another branch.
