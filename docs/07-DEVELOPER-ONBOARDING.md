# Developer Onboarding Guide

This guide helps new developers get started with the Outlook Add-in project.

## Prerequisites

### Required Software
- **Node.js** v16+ (recommended: v18 or v20)
- **npm** (comes with Node.js)
- **Git** for version control
- **VS Code** (recommended) or your preferred IDE
- **Outlook** (Web, Mac, or Windows) for testing

### Required Knowledge
- TypeScript/JavaScript
- React (hooks, components, state management)
- Office.js API (basic understanding)
- Azure AD / MSAL (authentication concepts)
- REST APIs

## Initial Setup

### 1. Clone Repository
```bash
git clone <repository-url>
cd WIP-AddIn
```

### 2. Install Dependencies
```bash
npm install
```

This installs all required packages including:
- React and React DOM
- TypeScript
- Fluent UI components
- MSAL libraries
- Office.js types
- Webpack and build tools

### 3. Set Up HTTPS Certificate
Required for Outlook Web and Mac development:

```bash
npx office-addin-dev-certs install
```

This generates and trusts a local HTTPS certificate.

### 4. Configure Environment
Copy and configure `environments.json`:

```bash
# Edit environments.json with your local configuration
# Update localhost environment with your Azure AD app registration details
```

Required configuration:
- Azure AD Client ID
- Azure AD Authority (tenant ID)
- Redirect URI (usually `https://localhost:3001`)
- API URLs (if available)
- API Keys (if available)

### 5. Start Development Server
```bash
npm start
```

The app will run at `https://localhost:3001`

### 6. Verify Setup
Open in browser:
- `https://localhost:3001/index.html` - Should show the app
- `https://localhost:3001/assets/icon-32.png` - Should show icon

If you see security warnings, proceed and trust the certificate.

## Project Structure Overview

### Key Directories
```
src/
├── components/     # React components
├── service/        # Business logic services
├── utils/          # Utility functions
├── config/         # Configuration
├── types/          # TypeScript types
├── auth/           # Authentication
└── hooks/          # Custom React hooks
```

### Key Files
- `src/index.tsx` - Application entry point
- `src/App.tsx` - Main app component
- `src/components/WorkbenchLanding.tsx` - Main UI
- `src/service/WorkbenchService.ts` - Main orchestration service
- `environments.json` - Runtime configuration

## Development Workflow

### 1. Making Changes
1. Make changes to source files in `src/`
2. Webpack will automatically rebuild
3. Browser will auto-reload (if configured)

### 2. Testing Changes
1. Sideload the add-in in Outlook (see below)
2. Test the functionality
3. Check browser console for errors
4. Use DebugService for logging

### 3. Debugging
- Open browser DevTools (F12)
- Check Console tab for DebugService logs
- Check Network tab for API calls
- Use breakpoints in Sources tab

See [Debugging Guide](./05-DEBUGGING.md) for details.

## Sideloading the Add-in

### Outlook Web
1. Open Outlook Web (https://outlook.office.com/)
2. Go to **Settings** (gear icon) → **View all Outlook settings**
3. Go to **Mail** → **General** → **Manage add-ins**
4. Click **+ Add a custom add-in** → **Add from file**
5. Select `Manifests/Manifest.Local.xml`
6. The add-in will appear in your Outlook ribbon

### Outlook Desktop (Windows/Mac)
1. Open Outlook
2. Go to **File** → **Manage Add-ins**
3. Click **+ Add a custom add-in** → **Add from file**
4. Select `Manifests/Manifest.Local.xml`
5. The add-in will appear in your Outlook ribbon

## Common Development Tasks

### Adding a New Service
1. Create file in `src/service/`
2. Follow singleton pattern:
   ```typescript
   class NewService {
     private static instance: NewService;
     private constructor() {}
     public static getInstance(): NewService {
       if (!NewService.instance) {
         NewService.instance = new NewService();
       }
       return NewService.instance;
     }
   }
   export default NewService.getInstance();
   ```
3. Use DebugService for logging
4. Add error handling

### Adding a New Component
1. Create file in `src/components/`
2. Use TypeScript and React hooks
3. Use Fluent UI components
4. Add props interface
5. Handle loading and error states

### Adding a New Utility Function
1. Create or add to file in `src/utils/`
2. Export function
3. Add TypeScript types
4. Add JSDoc comments

### Modifying Configuration
1. Edit `environments.json`
2. Add new environment if needed
3. Update URL patterns if needed
4. Restart dev server

## Code Style and Conventions

### TypeScript
- Use TypeScript for all new code
- Define interfaces for props and data structures
- Use type annotations
- Avoid `any` type

### React
- Use functional components with hooks
- Use TypeScript interfaces for props
- Handle loading and error states
- Use DebugService for logging

### Naming Conventions
- **Components:** PascalCase (e.g., `WorkbenchLanding`)
- **Services:** PascalCase (e.g., `WorkbenchService`)
- **Functions:** camelCase (e.g., `handleSubmit`)
- **Constants:** UPPER_SNAKE_CASE (e.g., `MAX_SIZE`)
- **Files:** Match export name (e.g., `WorkbenchService.ts`)

### File Organization
- One component/service per file
- Co-locate related files
- Use index files for exports if needed

## Testing

### Manual Testing
1. Test in Outlook Web
2. Test in Outlook Desktop (if available)
3. Test different email scenarios:
   - Sent emails
   - Draft emails
   - Emails with attachments
   - Emails without attachments

### Testing Checklist
- [ ] Authentication works
- [ ] Form submission works
- [ ] File validation works
- [ ] Error handling works
- [ ] Success/error dialogs appear
- [ ] Email forwarding works (if enabled)

## Debugging Tips

### Enable Debug Logging
Set in `environments.json`:
```json
{
  "DEBUG_ENABLED": true,
  "DEBUG_LEVEL": "debug"
}
```

### Common Issues

#### Port Already in Use
```bash
npm run kill-ports
npm start
```

#### SSL Certificate Issues
```bash
npx office-addin-dev-certs uninstall
npx office-addin-dev-certs install
```

#### Office.js Not Loading
- Check if Office.js is available: `typeof Office !== 'undefined'`
- Verify manifest is correctly sideloaded
- Check browser console for errors

#### Authentication Issues
- Check Azure AD configuration
- Verify redirect URI matches
- Check browser console for MSAL errors

## Documentation

### Available Documentation
1. **[Architecture Overview](./01-ARCHITECTURE.md)** - System architecture
2. **[Code Structure](./02-CODE-STRUCTURE.md)** - Code organization
3. **[Services Documentation](./03-SERVICES.md)** - Service details
4. **[Components Documentation](./04-COMPONENTS.md)** - Component details
5. **[Debugging Guide](./05-DEBUGGING.md)** - Debugging information
6. **[Configuration Guide](./06-CONFIGURATION.md)** - Configuration details

### Reading Code
1. Start with `src/index.tsx` (entry point)
2. Read `src/App.tsx` (main component)
3. Read `src/components/WorkbenchLanding.tsx` (main UI)
4. Read service files as needed
5. Check utility functions as needed

## Getting Help

### Resources
- **Documentation:** Check `docs/` folder
- **Code Comments:** Read inline comments
- **Console Logs:** Check DebugService output
- **Team:** Ask team members

### Information to Provide When Asking for Help
1. What you're trying to do
2. What error you're seeing
3. Browser console logs
4. Steps to reproduce
5. Environment (dev, local, etc.)

## Next Steps

### Learning Path
1. **Week 1:** Set up environment, understand project structure
2. **Week 2:** Learn services, understand data flow
3. **Week 3:** Learn components, understand UI flow
4. **Week 4:** Make small changes, test thoroughly

### Recommended Reading Order
1. Architecture Overview
2. Code Structure
3. Services Documentation
4. Components Documentation
5. Debugging Guide
6. Configuration Guide

### Practice Tasks
1. Add a new debug log message
2. Modify a component's styling
3. Add a new utility function
4. Modify error handling
5. Add a new configuration variable

## Best Practices

### Development
- ✅ Write clear, readable code
- ✅ Add comments for complex logic
- ✅ Use TypeScript types
- ✅ Handle errors gracefully
- ✅ Use DebugService for logging
- ✅ Test changes thoroughly

### Code Review
- ✅ Review your own code before committing
- ✅ Check for TypeScript errors
- ✅ Verify error handling
- ✅ Test in Outlook
- ✅ Check console for errors

### Git Workflow
- ✅ Create feature branches
- ✅ Write descriptive commit messages
- ✅ Test before committing
- ✅ Don't commit secrets
- ✅ Keep commits focused

## Troubleshooting

### Build Issues
```bash
# Clear node_modules and reinstall
rm -rf node_modules package-lock.json
npm install

# Clear webpack cache
rm -rf dist/
npm start
```

### TypeScript Errors
- Check `tsconfig.json` settings
- Verify type definitions are installed
- Check for missing imports

### Runtime Errors
- Check browser console
- Check DebugService logs
- Verify configuration
- Check network requests

## Additional Resources

### Office.js Documentation
- [Office Add-ins Documentation](https://docs.microsoft.com/office/dev/add-ins/)
- [Office.js API Reference](https://docs.microsoft.com/javascript/api/office)

### React Documentation
- [React Documentation](https://react.dev/)
- [React Hooks](https://react.dev/reference/react)

### Fluent UI Documentation
- [Fluent UI Documentation](https://developer.microsoft.com/en-us/fluentui)

### MSAL Documentation
- [MSAL.js Documentation](https://github.com/AzureAD/microsoft-authentication-library-for-js)

## Checklist for New Developers

### Setup
- [ ] Node.js installed
- [ ] Repository cloned
- [ ] Dependencies installed
- [ ] HTTPS certificate installed
- [ ] Environment configured
- [ ] Dev server running
- [ ] Add-in sideloaded in Outlook

### Understanding
- [ ] Read architecture overview
- [ ] Understand code structure
- [ ] Know main services
- [ ] Know main components
- [ ] Understand configuration

### Development
- [ ] Can make code changes
- [ ] Can test changes
- [ ] Can debug issues
- [ ] Can use DebugService
- [ ] Can read logs

### Ready to Contribute
- [ ] Can create new components
- [ ] Can create new services
- [ ] Can modify existing code
- [ ] Can test thoroughly
- [ ] Can debug effectively

Welcome to the team! 🎉

