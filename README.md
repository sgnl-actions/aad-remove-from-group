# Azure Active Directory Remove User from Group

This action removes a user from a group in Azure Active Directory using Microsoft Graph API. The action uses a two-step process to ensure reliable user identification and removal.

## Overview

The action performs the following steps:
1. **Get User**: Retrieves the user's directory object ID using their userPrincipalName (email)
2. **Remove from Group**: Removes the user from the specified group using their directory object ID

This two-step approach ensures that users are correctly identified even when their userPrincipalName contains special characters that require URL encoding.

## Prerequisites

- Azure AD application with appropriate permissions:
  - `User.Read.All` (to look up users)
  - `Group.ReadWrite.All` or `GroupMember.ReadWrite.All` (to modify group membership)
- Valid Azure AD access token

## Configuration

### Authentication

This action supports two OAuth2 authentication methods:

#### Option 1: OAuth2 Client Credentials
| Secret/Environment | Description |
|-------------------|-------------|
| `OAUTH2_CLIENT_CREDENTIALS_CLIENT_SECRET` | OAuth2 client secret |
| `OAUTH2_CLIENT_CREDENTIALS_CLIENT_ID` | OAuth2 client ID |
| `OAUTH2_CLIENT_CREDENTIALS_TOKEN_URL` | OAuth2 token endpoint URL |
| `OAUTH2_CLIENT_CREDENTIALS_SCOPE` | OAuth2 scope (optional) |
| `OAUTH2_CLIENT_CREDENTIALS_AUDIENCE` | OAuth2 audience (optional) |
| `OAUTH2_CLIENT_CREDENTIALS_AUTH_STYLE` | OAuth2 auth style (optional) |

#### Option 2: OAuth2 Authorization Code
| Secret | Description |
|--------|-------------|
| `OAUTH2_AUTHORIZATION_CODE_ACCESS_TOKEN` | OAuth2 access token |

### Required Environment Variables

| Variable | Description | Example |
|----------|-------------|---------|
| `ADDRESS` | Default Azure AD API base URL | `https://graph.microsoft.com` |

### Input Parameters

| Parameter | Type | Required | Description | Example |
|-----------|------|----------|-------------|---------|
| `userPrincipalName` | string | Yes | User Principal Name (email) of the user to remove | `user@example.com` |
| `groupId` | string | Yes | Azure AD Group ID (GUID) | `12345678-1234-1234-1234-123456789012` |
| `address` | string | No | Optional Azure AD API base URL override | `https://graph.microsoft.com` |

### Output Structure

| Field | Type | Description |
|-------|------|-------------|
| `status` | string | Operation result (success, failed, etc.) |
| `userPrincipalName` | string | User Principal Name that was processed |
| `groupId` | string | The group ID that was processed |
| `userId` | string | The Azure AD object ID of the user |
| `removed` | boolean | Whether the user was actually removed (true) or wasn't a member (false) |
| `address` | string | The Azure AD API base URL used |

## Development

### Local Testing

```bash
# Install dependencies
npm install

# Run unit tests
npm test

# Check test coverage (must be 80%+)
npm run test:coverage

# Run linting
npm run lint

# Build distribution
npm run build

# Test locally with sample parameters
npm run dev -- --params '{"userPrincipalName": "user@example.com", "groupId": "12345678-1234-1234-1234-123456789abc"}'
```

## Usage Examples

### Basic Usage

```json
{
  "userPrincipalName": "john.doe@company.com",
  "groupId": "12345678-1234-1234-1234-123456789abc"
}
```

### Special Characters in Email

The action handles userPrincipalNames with special characters correctly:

```json
{
  "userPrincipalName": "user+tag@company.com",
  "groupId": "12345678-1234-1234-1234-123456789abc"
}
```

### With OAuth2 Client Credentials

```json
{
  "script_inputs": {
    "userPrincipalName": "john.doe@company.com",
    "groupId": "12345678-1234-1234-1234-123456789abc",
    "address": "https://graph.microsoft.com"
  },
  "environment": {
    "ADDRESS": "https://graph.microsoft.com",
    "OAUTH2_CLIENT_CREDENTIALS_TOKEN_URL": "https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/token",
    "OAUTH2_CLIENT_CREDENTIALS_CLIENT_ID": "your-client-id",
    "OAUTH2_CLIENT_CREDENTIALS_SCOPE": "https://graph.microsoft.com/.default"
  },
  "secrets": {
    "OAUTH2_CLIENT_CREDENTIALS_CLIENT_SECRET": "your-client-secret"
  }
}
```

### With OAuth2 Authorization Code

```json
{
  "script_inputs": {
    "userPrincipalName": "john.doe@company.com",
    "groupId": "12345678-1234-1234-1234-123456789abc",
    "address": "https://graph.microsoft.com"
  },
  "environment": {
    "ADDRESS": "https://graph.microsoft.com"
  },
  "secrets": {
    "OAUTH2_AUTHORIZATION_CODE_ACCESS_TOKEN": "your-access-token"
  }
}
```

## API Endpoints

The action makes the following Microsoft Graph API calls:

1. **GET /users/{userPrincipalName}**
   - Retrieves user information including directory object ID
   - URL encodes the userPrincipalName parameter

2. **DELETE /groups/{groupId}/members/{userId}/$ref**
   - Removes the user from the group
   - URL encodes the userId parameter

## Error Handling

The action implements comprehensive error handling:

### Success Cases
- **204 No Content**: User successfully removed from group (`removed: true`)
- **404 Not Found**: User was not a member of the group (`removed: false`)

### Retryable Errors
- **429 Too Many Requests**: Rate limiting (waits 5 seconds before retry)
- **502/503/504**: Server errors (retried by framework)

### Fatal Errors
- **401 Unauthorized**: Invalid or expired token
- **403 Forbidden**: Insufficient permissions
- **400 Bad Request**: Invalid parameters

### Input Validation Errors
- Missing `userPrincipalName`
- Missing `groupId`

## Security Considerations

- **Token Security**: Never log or expose the Azure AD access token
- **URL Encoding**: All user inputs are properly URL encoded to prevent injection attacks
- **Least Privilege**: Use tokens with minimal required permissions
- **HTTPS Only**: All API calls use HTTPS for secure communication

## Troubleshooting

### Common Issues

1. **User not found during lookup**
   - Verify the userPrincipalName is correct
   - Ensure the token has `User.Read.All` permission

2. **Permission denied during group removal**
   - Verify the token has `Group.ReadWrite.All` or `GroupMember.ReadWrite.All` permission
   - Check that the group exists and the token has access to it

3. **Rate limiting (429 errors)**
   - The action automatically handles rate limits with exponential backoff
   - Consider reducing concurrent operations if needed

4. **URL encoding issues**
   - The action automatically handles URL encoding for special characters
   - No manual encoding is required in input parameters

## Testing

The test suite covers:
- Two-step process (user lookup + group removal)
- Success scenarios (204 and 404 responses)
- URL encoding for special characters
- Input validation
- Error handling for all scenarios
- Authentication and authorization errors

Run tests with:
```bash
npm test
npm run test:coverage  # Must achieve 80%+ coverage
```

## Deployment

1. **Run tests**: `npm test`
2. **Check coverage**: `npm run test:coverage`
3. **Lint code**: `npm run lint`
4. **Build distribution**: `npm run build`
5. **Tag release**: `git tag v1.0.0 -m "Initial release"`
6. **Push to GitHub**: `git push origin v1.0.0`

## Usage in SGNL

```json
{
  "job_request": {
    "name": "remove-user-from-group",
    "type": "nodejs-22",
    "script": {
      "repository": "github.com/sgnl-actions/aad-remove-from-group@main",
      "type": "nodejs"
    },
    "script_inputs": {
      "userPrincipalName": "user@company.com",
      "groupId": "12345678-1234-1234-1234-123456789abc",
      "address": "https://graph.microsoft.com"
    },
    "environment": {
      "ADDRESS": "https://graph.microsoft.com",
      "OAUTH2_CLIENT_CREDENTIALS_TOKEN_URL": "https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/token",
      "OAUTH2_CLIENT_CREDENTIALS_CLIENT_ID": "your-client-id",
      "OAUTH2_CLIENT_CREDENTIALS_SCOPE": "https://graph.microsoft.com/.default"
    },
    "secrets": {
      "OAUTH2_CLIENT_CREDENTIALS_CLIENT_SECRET": "your-client-secret"
    }
  }
}
```