# Components Documentation

This document provides detailed information about all React components in the application.

## Component Overview

Components are located in `src/components/` directory and use Fluent UI for styling and UI elements.

## Component Hierarchy

```
App
  └── WorkbenchLanding
      ├── WorkbenchHeader
      ├── LandingSection
      ├── BUProductsSection
      ├── WorkbenchDialogs
      │   ├── SuccessMessage
      │   ├── ErrorMessage
      │   ├── ConfirmationDialog
      │   └── RetryButton
      └── SpinnerOverlay
```

## Main Components

### 1. App Component
**File:** `src/App.tsx`  
**Lines:** ~222  
**Purpose:** Main application component that handles authentication and renders the appropriate UI.

#### Responsibilities
- Authentication state management
- Token acquisition and refresh
- Authentication UI rendering
- Main UI rendering (after authentication)

#### State Management
```typescript
const [isAuthenticated, setIsAuthenticated] = useState<boolean | null>(null);
const [apiToken, setApiToken] = useState<string | null>(null);
const [graphToken, setGraphToken] = useState<string | null>(null);
const [authError, setAuthError] = useState<string | null>(null);
const [isRetrying, setIsRetrying] = useState(false);
const [isInitializing, setIsInitializing] = useState(true);
```

#### Key Features
- **Token Auto-Refresh:** Refreshes tokens every 5 minutes
- **Initial Auth Check:** Checks for cached tokens on mount
- **Error Handling:** Shows error messages and retry button
- **Loading States:** Shows spinner during initialization and authentication

#### Render States
1. **Initializing:** Shows spinner while checking authentication
2. **Authenticating:** Shows spinner during authentication
3. **Not Authenticated:** Shows sign-in button or error message
4. **Authenticated:** Shows WorkbenchLanding component

#### Methods
- `authenticate()` - Initiates authentication flow
- `refreshTokens()` - Refreshes both tokens
- `handleRetry()` - Retries authentication after error

---

### 2. WorkbenchLanding Component
**File:** `src/components/WorkbenchLanding.tsx`  
**Lines:** ~581  
**Purpose:** Main UI component for the workbench submission flow.

#### Responsibilities
- Form handling (BU, Product selection)
- File validation UI
- Email metadata display
- Submission orchestration
- Success/error dialog management
- Duplicate detection

#### Props
```typescript
interface WorkbenchLandingProps {
  apiToken: string | null;
  graphToken: string | null;
}
```

#### State Management (15+ useState hooks)
```typescript
// UI State
const [showLanding, setShowLanding] = useState(true);
const [showBUProducts, setShowBUProducts] = useState(false);
const [showSuccessMessage, setShowSuccessMessage] = useState(false);
const [showFailureMessage, setShowFailureMessage] = useState(false);
const [showLoadingMessage, setShowLoadingMessage] = useState(false);
const [showConfirmationDialog, setShowConfirmationDialog] = useState(false);

// Form State
const [selectedProduct, setSelectedProduct] = useLocalStorage(...);
const [selectedBU, setSelectedBU] = useLocalStorage(...);
const [sendCopyToCyberAdmin, setSendCopyToCyberAdmin] = useLocalStorage(...);

// Error State
const [forwardingFailed, setForwardingFailed] = useState(false);
const [forwardingFailedReason, setForwardingFailedReason] = useState<string>();

// Validation State
const [fileValidationResult, setFileValidationResult] = useState<FileValidationResult>(...);
const [isValidatingFiles, setIsValidatingFiles] = useState(false);
const [isDuplicate, setIsDuplicate] = useState(false);

// Email State
const [emailSubject, setEmailSubject] = useState("");
const [isDraftEmail, setIsDraftEmail] = useState(false);

// Retry State
const [lastPlacementId, setLastPlacementId] = useState<string>();
const [lastGraphItemId, setLastGraphItemId] = useState<string>();
const [lastSharedMailbox, setLastSharedMailbox] = useState<string>();
```

#### Key Features
1. **Multi-Step Form:**
   - Landing page → BU/Product selection → Submission

2. **File Validation:**
   - Real-time validation on attachment changes
   - Shows validation errors
   - Prevents submission if invalid

3. **Duplicate Detection:**
   - Checks for duplicate submissions
   - Shows warning if duplicate found

4. **Email Metadata:**
   - Displays email subject
   - Shows draft email indicator
   - Extracts email information

5. **Submission Flow:**
   - Validates files
   - Checks for duplicates
   - Submits to WorkbenchService
   - Shows success/error dialogs

#### Main Methods

##### `handleSubmit()`
Handles form submission.

**Process:**
1. Validates files
2. Checks for duplicates
3. Shows confirmation dialog
4. Submits to WorkbenchService
5. Handles success/error

##### `handleDownloadEmail()`
Downloads email as EML file (for testing).

##### `handleRetryForward()`
Retries email forwarding after failure.

#### UI Sections
1. **Landing Section:** Initial welcome screen
2. **BU Products Section:** Form for BU and Product selection
3. **Header:** Navigation and title
4. **Dialogs:** Success, error, confirmation dialogs

---

### 3. WorkbenchDialogs Component
**File:** `src/components/WorkbenchDialogs.tsx`  
**Lines:** ~265  
**Purpose:** Dialog components for user feedback and confirmations.

#### Sub-Components

##### SuccessMessage
Shows success message after successful submission.

**Props:**
```typescript
{
  placementId: string;
  onClose: () => void;
}
```

**Features:**
- Displays placement ID
- Shows success icon
- Close button

##### ErrorMessage
Shows error message after failed submission.

**Props:**
```typescript
{
  error: string;
  onClose: () => void;
  onRetry?: () => void;
}
```

**Features:**
- Displays error message
- Shows error icon
- Retry button (optional)
- Close button

##### ConfirmationDialog
Shows confirmation dialog before submission.

**Props:**
```typescript
{
  isOpen: boolean;
  onConfirm: () => void;
  onCancel: () => void;
  emailSubject: string;
  isDraftEmail: boolean;
  fileValidationResult: FileValidationResult;
}
```

**Features:**
- Shows email subject
- Shows draft email warning
- Shows file validation status
- Confirm and Cancel buttons

##### RetryButton
Button for retrying failed operations.

**Props:**
```typescript
{
  onRetry: () => void;
  placementId?: string;
  graphItemId?: string;
  sharedMailbox?: string;
}
```

**Features:**
- Retry forwarding
- Shows retry information

---

### 4. WorkbenchHeader Component
**File:** `src/components/WorkbenchHeader.tsx`  
**Lines:** ~50  
**Purpose:** Header component with navigation and title.

#### Features
- Back button (navigates to landing)
- Title display
- Navigation between sections

#### Props
```typescript
{
  showBackButton: boolean;
  onBack: () => void;
  title: string;
}
```

---

### 5. BUProductsSection Component
**File:** `src/components/BUProductsSection.tsx`  
**Purpose:** Form section for Business Unit and Product selection.

#### Features
- Business Unit dropdown
- Product dropdown
- "Send copy to Cyber Admin" toggle
- Form validation
- Submit button

#### Props
```typescript
{
  selectedBU: string;
  selectedProduct: string;
  sendCopyToCyberAdmin: boolean;
  onBUChange: (bu: string) => void;
  onProductChange: (product: string) => void;
  onSendCopyChange: (value: boolean) => void;
  onSubmit: () => void;
  isSubmitDisabled: boolean;
}
```

#### Form Elements
- **BU Dropdown:** Selects business unit
- **Product Dropdown:** Selects product (depends on BU)
- **Toggle:** Enable/disable sending copy to Cyber Admin
- **Submit Button:** Submits form

---

### 6. LandingSection Component
**File:** `src/components/LandingSection.tsx`  
**Purpose:** Initial landing page section.

#### Features
- Welcome message
- Instructions
- "Get Started" button
- Navigation to main form

#### Props
```typescript
{
  onGetStarted: () => void;
}
```

---

### 7. SpinnerOverlay Component
**File:** `src/components/SpinnerOverlay.tsx`  
**Purpose:** Full-screen loading overlay.

#### Features
- Spinner display
- Loading message
- Blocks user interaction
- Centered layout

#### Props
```typescript
{
  isVisible: boolean;
  message?: string;
}
```

---

### 8. ErrorBoundary Component
**File:** `src/components/ErrorBoundary.tsx`  
**Purpose:** React error boundary for catching component errors.

#### Features
- Catches React component errors
- Shows error UI
- Prevents app crash
- Error logging

#### Implementation
```typescript
class ErrorBoundary extends React.Component {
  componentDidCatch(error: Error, errorInfo: React.ErrorInfo) {
    // Log error
  }
  
  render() {
    if (this.state.hasError) {
      return <ErrorUI />;
    }
    return this.props.children;
  }
}
```

---

## Component Communication

### Props Flow
```
App
  └── WorkbenchLanding (apiToken, graphToken)
      ├── WorkbenchHeader (showBackButton, onBack, title)
      ├── LandingSection (onGetStarted)
      ├── BUProductsSection (form props, onSubmit)
      └── WorkbenchDialogs (dialog props)
```

### State Management
- **Local State:** Each component manages its own UI state
- **Shared State:** WorkbenchLanding manages all submission state
- **Persistent State:** Uses `useLocalStorage` hook for form preferences

### Event Flow
```
User Action
    ↓
Component Handler
    ↓
Service Call
    ↓
State Update
    ↓
UI Re-render
```

## Styling

### Fluent UI Components Used
- `PrimaryButton`, `DefaultButton` - Buttons
- `Dropdown` - Dropdowns
- `Toggle` - Toggle switches
- `MessageBar` - Messages
- `Spinner` - Loading indicators
- `DatePicker` - Date selection
- `IconButton` - Icon buttons

### Custom Styles
- `SharedGrid.css` - Shared grid layout styles
- Inline styles for component-specific styling

## Component Best Practices

1. **Props Validation:** Use TypeScript interfaces for props
2. **Error Handling:** Wrap service calls in try-catch
3. **Loading States:** Show loading indicators during async operations
4. **User Feedback:** Show success/error messages
5. **Accessibility:** Use Fluent UI components (accessible by default)

## Common Patterns

### Form Handling Pattern
```typescript
const [value, setValue] = useState(initialValue);

const handleChange = (newValue: string) => {
  setValue(newValue);
};

return (
  <Dropdown
    value={value}
    onChange={(e, option) => handleChange(option?.key as string)}
  />
);
```

### Async Operation Pattern
```typescript
const [isLoading, setIsLoading] = useState(false);

const handleSubmit = async () => {
  setIsLoading(true);
  try {
    const result = await service.method();
    // Handle success
  } catch (error) {
    // Handle error
  } finally {
    setIsLoading(false);
  }
};
```

### Dialog Pattern
```typescript
const [isOpen, setIsOpen] = useState(false);

const handleOpen = () => setIsOpen(true);
const handleClose = () => setIsOpen(false);

return (
  <Dialog isOpen={isOpen} onDismiss={handleClose}>
    {/* Dialog content */}
  </Dialog>
);
```

## Component Testing Considerations

### Testable Aspects
- Component rendering
- Props handling
- User interactions
- State updates
- Error handling

### Mocking Requirements
- Service methods
- Office.js API
- MSAL authentication
- API responses

## Future Improvements

1. **State Management:** Consider useReducer for WorkbenchLanding
2. **Component Splitting:** Break down large components
3. **Custom Hooks:** Extract logic into reusable hooks
4. **Memoization:** Use React.memo for performance
5. **Error Boundaries:** More granular error boundaries

