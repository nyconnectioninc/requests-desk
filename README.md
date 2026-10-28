# Requests Desk

The private work dashboard for the NY Connection **Employee Requests** sheet. It is kept separate from the request form (`connection-request-hub`) that stores use.

- **Frontend:** `index.html`, served by GitHub Pages.
- **Backend:** `apps-script/Code.gs`, a standalone Google Apps Script web app that reads and writes the sheet.
- **Access:** sign in with an email code. Only emails listed as Active on the sheet's **Users** tab can get in. Sessions last 30 days per device.
- **Queues:** Shipping Labels, Business Cards, ID Badges, Shirts, Brinks Logins, Hire Requests. Store Supplies Orders is excluded on both the server and the page.

## What gets written to the sheet

| Column (added on first run) | Value |
|---|---|
| Status | blank (open) · In Progress · Needs Info · Done |
| Handled By | Name from the Users tab |
| Updated At | Time of the last change |
| Team Notes | Tracking #, what's missing, and so on |
| Repeat Of | Set by the request form when someone submits again while an earlier request is still open: the original's timestamp and row |
| Repeat Reason | Lost · Damaged · Name change · … or Not given |

## Repeat requests

Badges, Apparel, Business Cards and Brinks are checked before they're submitted. The form calls two
public actions on this API: `checkRepeat` (returns only yes/no, the date and status of the open request)
and `tagRepeat` (writes Repeat Of / Repeat Reason on the row that was just submitted, within 15 minutes,
once). People can always submit; the dashboard shows a **Repeat** pill so the team fills one and closes the other.

The **Activity Log** tab records every change as a row: time, who, which queue, which row, what changed.

## Setup (one time)

1. Go to **script.google.com**, choose **New project**, and name it `Requests Desk API`.
2. Paste in the contents of `apps-script/Code.gs`.
3. Select `setup` in the function menu and click **Run**. Approve the access prompts for Sheets and Gmail. This creates the **Users** and **Activity Log** tabs and adds `jon.sanchez@nyconnectionli.com` as Admin.
4. Choose **Deploy → New deployment → Web app**. Set Execute as **Me** and Who has access **Anyone**, then click **Deploy** and copy the Web app URL.
5. In `index.html`, set `API_URL` to that URL and commit.
6. In the GitHub repo, go to **Settings → Pages**, set the source to **Deploy from a branch**, pick `main` and `/ (root)`, and save.

## Updating

- **Page changes:** edit `index.html` and push. The live site updates in about a minute.
- **Backend changes:** edit `Code.gs` in Apps Script, then choose **Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy**. The URL stays the same.
- **Adding or removing people:** use **Team access** in the dashboard, or edit the **Users** tab directly.
