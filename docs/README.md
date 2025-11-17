# Documentation Index

Welcome to the Outlook Add-in documentation. This folder contains comprehensive documentation for developers, support staff, and future maintainers.

## Documentation Structure

### 1. [Architecture Overview](./01-ARCHITECTURE.md)
High-level system architecture, technology stack, design patterns, and integration points.

**Read this first** to understand the overall system design.

### 2. [Code Structure](./02-CODE-STRUCTURE.md)
Detailed code organization, directory structure, file purposes, and code metrics.

**Use this** to navigate the codebase and understand file organization.

### 3. [Services Documentation](./03-SERVICES.md)
Complete documentation of all services, their purposes, methods, and usage patterns.

**Reference this** when working with services or understanding business logic.

### 4. [Components Documentation](./04-COMPONENTS.md)
Detailed documentation of all React components, their props, state, and responsibilities.

**Reference this** when working with UI components.

### 5. [Debugging Guide](./05-DEBUGGING.md)
Comprehensive debugging information, common issues, troubleshooting steps, and debugging tools.

**Use this** when debugging issues or investigating problems.

### 6. [Configuration Guide](./06-CONFIGURATION.md)
Complete configuration documentation, environment setup, and configuration variables.

**Reference this** when setting up environments or modifying configuration.

### 7. [Developer Onboarding Guide](./07-DEVELOPER-ONBOARDING.md)
Step-by-step guide for new developers to get started with the project.

**Start here** if you're new to the project.

## Quick Reference

### For New Developers
1. Start with [Developer Onboarding Guide](./07-DEVELOPER-ONBOARDING.md)
2. Read [Architecture Overview](./01-ARCHITECTURE.md)
3. Review [Code Structure](./02-CODE-STRUCTURE.md)
4. Reference [Services](./03-SERVICES.md) and [Components](./04-COMPONENTS.md) as needed

### For Support Staff
1. Review [Architecture Overview](./01-ARCHITECTURE.md) for system understanding
2. Use [Debugging Guide](./05-DEBUGGING.md) for troubleshooting
3. Reference [Configuration Guide](./06-CONFIGURATION.md) for environment issues
4. Check [Services Documentation](./03-SERVICES.md) for service-specific issues

### For Developers Working on Features
1. Review [Code Structure](./02-CODE-STRUCTURE.md) to find relevant files
2. Reference [Services Documentation](./03-SERVICES.md) for service usage
3. Reference [Components Documentation](./04-COMPONENTS.md) for component patterns
4. Use [Debugging Guide](./05-DEBUGGING.md) for testing and debugging

### For Code Reviewers
1. Understand [Architecture Overview](./01-ARCHITECTURE.md) for design context
2. Reference [Code Structure](./02-CODE-STRUCTURE.md) for organization patterns
3. Check [Services](./03-SERVICES.md) and [Components](./04-COMPONENTS.md) for implementation patterns

## Key Concepts

### Application Flow
1. **Initialization:** Office.js → Runtime Config → MSAL → App
2. **Authentication:** MSAL → Token Acquisition → Token Caching
3. **Submission:** File Validation → Email Conversion → API Submission → Email Forwarding

### Service Pattern
All services use singleton pattern:
```typescript
const service = ServiceName.getInstance();
const result = await service.method();
```

### Component Pattern
Components use React hooks for state:
```typescript
const [state, setState] = useState(initialValue);
const handleAction = async () => { /* ... */ };
```

### Configuration Pattern
Configuration loaded at runtime:
```typescript
await runtimeConfig.initialize();
const value = runtimeConfig.getString('KEY');
```

## Common Tasks

### Setting Up Development Environment
See [Developer Onboarding Guide](./07-DEVELOPER-ONBOARDING.md)

### Debugging an Issue
See [Debugging Guide](./05-DEBUGGING.md)

### Understanding a Service
See [Services Documentation](./03-SERVICES.md)

### Understanding a Component
See [Components Documentation](./04-COMPONENTS.md)

### Configuring an Environment
See [Configuration Guide](./06-CONFIGURATION.md)

## Documentation Maintenance

### When to Update Documentation
- When adding new services
- When adding new components
- When changing architecture
- When adding new configuration
- When fixing common issues

### Documentation Standards
- Keep documentation up to date
- Include code examples
- Explain "why" not just "what"
- Include troubleshooting information
- Keep it concise but complete

## Additional Resources

### External Documentation
- [Office Add-ins Documentation](https://docs.microsoft.com/office/dev/add-ins/)
- [React Documentation](https://react.dev/)
- [Fluent UI Documentation](https://developer.microsoft.com/en-us/fluentui)
- [MSAL.js Documentation](https://github.com/AzureAD/microsoft-authentication-library-for-js)

### Project Files
- `README.md` - Project overview
- `package.json` - Dependencies and scripts
- `tsconfig.json` - TypeScript configuration
- `webpack.config.js` - Build configuration
- `environments.json` - Runtime configuration

## Getting Help

### Documentation Issues
If you find errors or missing information in the documentation:
1. Check if the information exists in another document
2. Update the documentation if you know the answer
3. Ask the team for clarification

### Code Issues
If you have questions about the code:
1. Check relevant documentation
2. Review code comments
3. Check DebugService logs
4. Ask the team

## Document Version

**Last Updated:** 2025-11-11  
**Version:** 1.0  
**Maintained By:** Development Team

---

**Note:** This documentation is a living document. It should be updated as the codebase evolves. If you make significant changes to the codebase, please update the relevant documentation.

