# KPC Microsoft Sign-in Test

A minimal single-tenant Microsoft Entra ID sign-in test for the Kamloops Pickleball Club. It is intended to be tested with the existing Microsoft 365 accounts for Smash and Erne.

## Test accounts

- `smash@kamloopspickleballclub.onmicrosoft.com`
- `erne@kamloopspickleballclub.onmicrosoft.com`

The app never receives or stores their Microsoft 365 passwords.

## Microsoft Entra setup

1. Open **Microsoft Entra admin centre**.
2. Go to **Identity > Applications > App registrations**.
3. Select **New registration**.
4. Name it `KPC Sign-in Test`.
5. Choose **Accounts in this organizational directory only**.
6. Register the application.
7. Copy the **Application (client) ID** and **Directory (tenant) ID**.
8. Open **Authentication > Add a platform > Single-page application**.
9. For local testing, add `http://localhost:5173`.
10. After Netlify deployment, add the exact Netlify site URL, such as `https://your-site-name.netlify.app`.

Do not create a client secret. A browser-based single-page application uses a public client configuration.

## Add the two IDs

Edit `src/config.js`:

```js
export const entraConfig = {
  clientId: "YOUR_APPLICATION_CLIENT_ID",
  tenantId: "YOUR_DIRECTORY_TENANT_ID",
};
```

The client ID and tenant ID identify the app and tenant. They are not passwords.

## Run locally

```bash
npm install
npm run dev
```

Then open `http://localhost:5173`.

## Deploy through Netlify

1. In Netlify, choose **Add new site > Import an existing project**.
2. Connect `ITKPC/kpc-msft-signin`.
3. Netlify will use `netlify.toml` and run `npm run build`.
4. Copy the final Netlify URL.
5. Add that exact URL as a **Single-page application redirect URI** in the Entra app registration.
6. Redeploy after changing the IDs in `src/config.js`.

## Security boundary

This demonstration confirms that Microsoft authenticated the account and conditionally displays the signed-in screen. Do not put confidential member information, secrets, or privileged operations directly in this front-end. Those require a protected server or API that validates Microsoft access tokens before returning private data.
