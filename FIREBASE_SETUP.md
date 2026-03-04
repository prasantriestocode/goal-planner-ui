# Firebase Setup

## 1. Create project
1. Open https://console.firebase.google.com
2. Create project (or reuse existing).
3. Enable Authentication -> Sign-in method -> Email/Password.
4. Create Firestore database (Production mode or Test mode for initial setup).

## 2. Web app config
1. Project Settings -> General -> Your apps -> Web app.
2. Copy SDK config values.
3. Paste values into `firebase-config.js`.

## 3. Firestore structure used by app
- `users/{uid}`
  - `email`
  - `role` (`admin` or `investor`)
  - `investorName`
- `investorPlans/{uid}`
  - `investorName`
  - `model`
  - `goals`
  - `additionalProperties`
  - `networthNotes`
  - `adminPortfolio`
  - `updatedAt`

## 4. Firestore rules (single admin email)
Replace rules with the block below. Update `ops@aarthashastra.com` if needed.

```txt
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    function isSignedIn() { return request.auth != null; }
    function isOwner(uid) { return isSignedIn() && request.auth.uid == uid; }
    function isAdmin() { return isSignedIn() && request.auth.token.email == "ops@aarthashastra.com"; }

    match /users/{uid} {
      allow read, write: if isOwner(uid) || isAdmin();
    }

    match /investorPlans/{uid} {
      allow read: if isOwner(uid) || isAdmin();
      allow write: if isOwner(uid) || isAdmin();
    }
  }
}
```

## 5. Deploy
- Push to GitHub `main`.
- Vercel auto redeploys from GitHub.
