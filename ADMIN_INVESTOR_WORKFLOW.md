# Admin Investor Management Workflow

## Overview

This document describes the **Phase 3: Admin Investor Management** implementation for the Goal Planner application. The workflow has been redesigned to enforce admin-only investor creation, with investors unable to self-signup.

## Implementation Summary

### 1. Admin Investor Creation

**Feature:** Admins can create investors directly through the admin panel form.

**UI Location:** Admin Panel → "➕ Add New Investor" section
- **Email**: Investor's email address
- **Investor Name**: Display name for the investor
- **Temporary Password**: Initial password for login

**Backend Function:** `createInvestor(email, investorName, password)`
- Creates Firebase Auth user with provided credentials
- Writes to `users/{uid}` collection with role="investor"
- Creates empty `investorPlans/{uid}` document with default structure
- Returns success message with shareable credentials

**Workflow:**
1. Admin fills email, investor name, and temporary password
2. Admin clicks "Create Investor" button
3. System creates Auth user and Firestore documents
4. Success message displays password for admin to share
5. Investor dropdown refreshes to show new investor
6. Message auto-hides after 5 seconds

### 2. Restricted Self-Signup

**Feature:** Only the admin email can self-signup; investors cannot create their own accounts.

**Modified Function:** `signup()`
- Checks if email matches admin email (ops@aarthashastra.com)
- Non-admin signup attempts show error: "Investor accounts can only be created by admin. Please contact support."
- Only admin email can successfully signup

**Why:** Ensures all investors are created and managed by the admin, maintaining data integrity and control.

### 3. My Page Locked After Questionnaire

**Feature:** After an investor completes the questionnaire, My Page becomes read-only.

**Locked Inputs:**
- Inflation Rates (General, Education, Marriage)
- Return Rates (Pre-Retirement, Post-Retirement, Savings Growth, Debt)
- Retirement Planning (Age, Life Expectancy, Monthly Expense)

**Implementation:**
- `lockMyPage()` function disables all My Page inputs
- Adds green "🔒 Locked" banner to indicate read-only status
- Set disabled=true and title attribute on inputs
- Called in `wizardFinish()` after questionnaire submission
- Also called when loading plan if `model.wizardCompleted` is true

**Why:** Prevents investors from changing assumptions after questionnaire submission, ensuring plan consistency.

### 4. Admin Password Reset

**Feature:** Admins can reset an investor's password securely.

**UI Location:** Admin Panel → Investor Selection → "🔐 Reset Password" button

**Workflow:**
1. Admin selects investor from dropdown
2. Admin clicks "Reset Password" button
3. System sends password reset email to investor's address
4. Investor receives email and clicks link to set new password
5. Success message shows in admin panel

**Backend Function:** `resetInvestorPassword(investorUid, email, newPassword)`
- Sends Firebase password reset email
- Investor self-serves password change via email link
- Most secure approach (no temporary passwords shared)

**Why:** Uses Firebase's built-in email verification for security.

---

## Architecture

### Data Structure

```
users/{uid}
├── email (string)
├── role (enum: "admin" | "investor")
├── investorName (string)
└── createdAt (timestamp)

investorPlans/{uid}
├── investorName (string)
├── model (object) - Form values from questionnaire
├── goals (array)
├── additionalProperties (array)
├── netlowerthNotes (string)
├── adminPortfolio (object)
├── customAssets (object)
├── customLiabilities (array)
├── lifeInsuranceRows (array)
├── healthInsuranceRows (array)
├── carInsuranceRows (array)
├── propertyInsuranceRows (array)
├── customExpenses (array)
├── cashflowOverrides (object)
├── children (array)
├── assetGrowthRates (object)
└── updatedAt (timestamp)
```

### Firestore Security Rules

**Location:** `firestore.rules`

**Rules Summary:**
- **Admin Email:** ops@aarthashastra.com can read/write all documents
- **Investors:** Can only read/write their own investorPlans document
- **User Creation:** Only admin can create new users (prevents self-signup)
- **Helper Functions:**
  - `isAdmin()`: Checks if current user email is admin email
  - `isOwner(userId)`: Checks if current user UID matches document owner

**Deployment:**
```bash
firebase deploy --only firestore:rules
```

---

## Testing Checklist

### Admin Features

- [ ] **Create Investor**
  - [ ] Admin can access "Add New Investor" form
  - [ ] Form validates all fields required
  - [ ] Investor successfully created in Firebase Auth
  - [ ] Investor appears in dropdown with (ID suffix)
  - [ ] Success message shows password
  - [ ] Form clears after success
  - [ ] Error handling for duplicate email shows proper message

- [ ] **View/Switch Investor**
  - [ ] Admin can select investor from dropdown
  - [ ] Plan data loads correctly for selected investor
  - [ ] "Investor Name" field shows selected investor
  - [ ] Loading skeleton appears briefly during switch

- [ ] **Reset Password**
  - [ ] Admin can click "Reset Password" button when investor selected
  - [ ] Password reset email sent to investor email
  - [ ] Success message displays investor email
  - [ ] Message auto-hides after 5 seconds
  - [ ] Error handling shows if investor not found

### Investor Features

- [ ] **Self-Signup Blocked**
  - [ ] Non-admin email signup attempt shows error message
  - [ ] Investor cannot create own account
  - [ ] "Create Account" button still visible (for admin use)
  - [ ] Error message instructs to contact admin

- [ ] **Login with Admin Credentials**
  - [ ] Investor receives email and password from admin
  - [ ] Investor can login with those credentials
  - [ ] Dashboard loads after login
  - [ ] Investor data persists correctly

- [ ] **Questionnaire & My Page**
  - [ ] Investor can access Questionnaire button
  - [ ] Wizard opens on first login
  - [ ] Investor can complete all 7 steps
  - [ ] After finishing, My Page loads and is read-only
  - [ ] "🔒 Locked" banner appears on My Page
  - [ ] All My Page inputs are disabled (cannot type)
  - [ ] Reload preserves locked state

- [ ] **Data Visibility**
  - [ ] Investor can view Dashboard
  - [ ] Investor can view Goal-Sheet
  - [ ] Investor can view Networth Statement
  - [ ] Investor can view Cash Flow
  - [ ] Investor can view My Page (read-only after questionnaire)
  - [ ] Cannot edit My Page assumptions after questionnaire

### Security & Permissions

- [ ] **Firestore Rules Enforcement**
  - [ ] Investor cannot read/write other investors' plans
  - [ ] Admin can switch between any investor
  - [ ] Investor cannot edit other investors' data
  - [ ] Users collection properly restricted

- [ ] **Auth Flow**
  - [ ] Admin email required for self-signup
  - [ ] Investor role assigned automatically to created users
  - [ ] Admin role assigned only to admin email
  - [ ] Password reset emails sent only to valid investor emails

### Data Persistence

- [ ] **Save & Load**
  - [ ] Investor data saves correctly to Firestore
  - [ ] Data persists after page reload
  - [ ] Data persists after logout/login
  - [ ] Admin can view investor data after switching

- [ ] **Export Functions**
  - [ ] Excel export includes all investor data
  - [ ] PDF export renders correctly
  - [ ] Exports include correct investor name

---

## Deployment Instructions

### 1. Verify Code Changes

```bash
cd /path/to/goal-planner
git log --oneline | head -5
# Should show:
# - "Add Firestore security rules for role-based access control"
# - "Implement admin password reset functionality"
# - "Restrict signup to admin only and lock My Page after questionnaire"
# - "Implement admin investor creation workflow"
```

### 2. Push to GitHub

```bash
git push origin main
```

### 3. Verify Vercel Deployment

- GitHub push triggers automatic Vercel deployment
- Check Vercel dashboard for deployment status
- Verify production URL is publicly accessible

### 4. Deploy Firestore Rules

```bash
# Install Firebase CLI (if not already installed)
npm install -g firebase-tools

# Login to Firebase
firebase login

# Deploy rules
firebase deploy --only firestore:rules

# Verify in Firebase Console > Firestore > Rules
```

### 5. Test in Production

- Visit production URL
- Test admin login with ops@aarthashastra.com
- Create a test investor
- Test investor login with provided credentials
- Verify My Page is locked after questionnaire
- Test password reset

---

## Known Limitations & Future Enhancements

### Current Limitations

1. **Password Reset Flow**
   - Uses email-based reset (investor must check email)
   - No immediate password change in admin panel
   - Investor must complete email verification

2. **Admin User Management**
   - Only one admin email (ops@aarthashastra.com) supported
   - Cannot add multiple admins through UI
   - Would require code changes to support multiple admins

3. **Questionnaire Lock**
   - My Page locked permanently (no unlock option)
   - Admin cannot edit investor assumptions
   - Would require additional UI/permissions for admin edit

### Future Enhancements

1. **Multi-Admin Support**
   - Store admin emails in Firestore admin collection
   - Update Firestore rules to check admin collection
   - Add admin management UI

2. **Flexible My Page Permissions**
   - Add checkbox: "Admin can edit investor assumptions"
   - Allow admin to override My Page lock
   - Add audit log of who changed what

3. **Investor Lifecycle Management**
   - Delete investor (with cascade delete)
   - Archive inactive investors
   - Bulk import investors from CSV

4. **Improved Password Reset**
   - Alternative: temporary password shown in admin panel
   - Option to generate random passwords
   - Encrypted password storage for sharing

---

## Troubleshooting

### Issue: "Firebase not initialized" error

**Solution:** Wait for Firebase to initialize before attempting operations. All Firebase operations have initialization checks.

### Issue: Investor not appearing in dropdown

**Solution:**
1. Verify investor was created successfully (check success message)
2. Try refreshing the page
3. Check Firestore console to confirm document exists

### Issue: My Page not locking after questionnaire

**Solution:**
1. Verify `model.wizardCompleted` is set to true
2. Clear browser cache and reload
3. Check browser console for errors
4. Verify `lockMyPage()` function is being called

### Issue: Password reset email not received

**Solution:**
1. Check investor email is correct in users collection
2. Verify email address is valid
3. Check spam/junk folder
4. Verify Firebase email configuration (check Firebase Console)
5. Try using different test email

### Issue: Firestore rules rejected my requests

**Solution:**
1. Verify current user is logged in and has correct role
2. Check `currentUser` in browser console
3. Verify Firestore rules are deployed correctly
4. Review Firebase error message for specific denied rule
5. Temporarily set rules to allow for debugging (not for production)

---

## Files Modified

### HTML (`index.html`)
- Added "Add New Investor" form section (lines 107-124)
- Added "Reset Password" button and message div (after line 104)

### JavaScript (`app.js`)
- Modified `signup()` to restrict to admin only (line 359)
- Added `createInvestor()` function (line 403)
- Added event listener for Create Investor button (line 738)
- Added `lockMyPage()` function (line 2327)
- Added event listener for Reset Password button (line 801)
- Modified `loadPlan()` to lock My Page if completed (line 483)
- Added `resetInvestorPassword()` function (line 469)

### CSS (`styles.css`)
- Added `.admin-divider` styling (line 413)
- Added `.admin-grid.three` for three-column layout (line 416)
- Added responsive mobile layouts

### New Files
- `firestore.rules`: Security rules for Firestore
- `ADMIN_INVESTOR_WORKFLOW.md`: This documentation file

---

## Commits Created

1. **ef071f7** - Implement admin investor creation workflow
2. **420b75d** - Restrict signup to admin only and lock My Page
3. **9bc4ed4** - Implement admin password reset functionality
4. **71d25c9** - Add Firestore security rules

---

## Contact & Support

For issues or questions about the admin investor workflow:
1. Review the troubleshooting section above
2. Check browser console for error messages
3. Verify Firestore rules are deployed
4. Review Firebase error messages in console
