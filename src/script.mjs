/**
 * Azure Active Directory Remove User from Group Action
 *
 * This action removes a user from a group in Azure Active Directory using a two-step process:
 * 1. Get the user's directory object ID using their userPrincipalName
 * 2. Remove the user from the group using the directory object ID
 */

/**
 * Get OAuth2 access token using client credentials flow
 * @param {Object} config - OAuth2 configuration
 * @returns {Promise<string>} Access token
 */
async function getClientCredentialsToken(config) {
  const { tokenUrl, clientId, clientSecret, scope, audience, authStyle } = config;

  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');

  if (scope) {
    params.append('scope', scope);
  }

  if (audience) {
    params.append('audience', audience);
  }

  const headers = {
    'Content-Type': 'application/x-www-form-urlencoded',
    'Accept': 'application/json'
  };

  if (authStyle === 'InParams') {
    params.append('client_id', clientId);
    params.append('client_secret', clientSecret);
  } else {
    const credentials = Buffer.from(`${clientId}:${clientSecret}`).toString('base64');
    headers['Authorization'] = `Basic ${credentials}`;
  }

  const response = await fetch(tokenUrl, {
    method: 'POST',
    headers,
    body: params.toString()
  });

  if (!response.ok) {
    let errorText;
    try {
      const errorData = await response.json();
      errorText = JSON.stringify(errorData);
    } catch {
      errorText = await response.text();
    }
    throw new Error(
      `OAuth2 token request failed: ${response.status} ${response.statusText} - ${errorText}`
    );
  }

  const data = await response.json();

  if (!data.access_token) {
    throw new Error('No access_token in OAuth2 response');
  }

  return data.access_token;
}

/**
 * Helper function to get a user by userPrincipalName
 * @param {string} userPrincipalName - User Principal Name (email)
 * @param {string} address - Azure AD base URL
 * @param {string} accessToken - OAuth2 access token
 * @returns {Promise<Response>} HTTP response
 */
async function getUser(userPrincipalName, address, accessToken) {
  // Remove trailing slash from address if present
  const cleanAddress = address.endsWith('/') ? address.slice(0, -1) : address;

  const encodedUPN = encodeURIComponent(userPrincipalName);
  const url = `${cleanAddress}/v1.0/users/${encodedUPN}`;

  const authHeader = accessToken.startsWith('Bearer ') ? accessToken : `Bearer ${accessToken}`;

  const response = await fetch(url, {
    method: 'GET',
    headers: {
      'Authorization': authHeader,
      'Accept': 'application/json'
    }
  });

  return response;
}

/**
 * Helper function to remove a user from a group
 * @param {string} groupId - Azure AD Group ID
 * @param {string} userId - User's directory object ID
 * @param {string} address - Azure AD base URL
 * @param {string} accessToken - OAuth2 access token
 * @returns {Promise<Response>} HTTP response
 */
async function removeUserFromGroup(groupId, userId, address, accessToken) {
  // Remove trailing slash from address if present
  const cleanAddress = address.endsWith('/') ? address.slice(0, -1) : address;

  const encodedUserId = encodeURIComponent(userId);
  const url = `${cleanAddress}/v1.0/groups/${groupId}/members/${encodedUserId}/$ref`;

  const authHeader = accessToken.startsWith('Bearer ') ? accessToken : `Bearer ${accessToken}`;

  const response = await fetch(url, {
    method: 'DELETE',
    headers: {
      'Authorization': authHeader,
      'Accept': 'application/json'
    }
  });

  return response;
}

export default {
  /**
   * Main execution handler - removes a user from an Azure AD group
   * @param {Object} params - Job input parameters
   * @param {string} params.userPrincipalName - User Principal Name (email) to remove from group
   * @param {string} params.groupId - Azure AD Group ID to remove user from
   * @param {string} params.address - The Azure AD API base URL (e.g., https://graph.microsoft.com)
   * @param {Object} context - Execution context with env, secrets, outputs
   * @param {string} context.environment.ADDRESS - Default Azure AD API base URL
   *
   * The configured auth type will determine which of the following environment variables and secrets are available
   * @param {string} context.secrets.OAUTH2_CLIENT_CREDENTIALS_CLIENT_SECRET
   * @param {string} context.environment.OAUTH2_CLIENT_CREDENTIALS_AUDIENCE
   * @param {string} context.environment.OAUTH2_CLIENT_CREDENTIALS_AUTH_STYLE
   * @param {string} context.environment.OAUTH2_CLIENT_CREDENTIALS_CLIENT_ID
   * @param {string} context.environment.OAUTH2_CLIENT_CREDENTIALS_SCOPE
   * @param {string} context.environment.OAUTH2_CLIENT_CREDENTIALS_TOKEN_URL
   *
   * @param {string} context.secrets.OAUTH2_AUTHORIZATION_CODE_ACCESS_TOKEN
   * @param {string} context.secrets.OAUTH2_AUTHORIZATION_CODE_AUTHORIZATION_CODE
   * @param {string} context.secrets.OAUTH2_AUTHORIZATION_CODE_CLIENT_SECRET
   * @param {string} context.secrets.OAUTH2_AUTHORIZATION_CODE_REFRESH_TOKEN
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_AUTH_STYLE
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_AUTH_URL
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_CLIENT_ID
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_LAST_TOKEN_ROTATION_TIMESTAMP
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_REDIRECT_URI
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_SCOPE
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_TOKEN_LIFETIME_FREQUENCY
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_TOKEN_ROTATION_FREQUENCY
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_TOKEN_ROTATION_INTERVAL
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_TOKEN_URL
   *
   * @returns {Object} Job results
   */
  invoke: async (params, context) => {
    console.log('Starting Azure AD remove user from group action');

    // Validate inputs
    if (!params.userPrincipalName) {
      throw new Error('userPrincipalName is required');
    }

    if (!params.groupId) {
      throw new Error('groupId is required');
    }

    // Determine the URL to use
    const address = params.address || context.environment?.ADDRESS;
    if (!address) {
      throw new Error('No URL specified. Provide either address parameter or ADDRESS environment variable');
    }

    let accessToken;

    if (context.secrets?.OAUTH2_AUTHORIZATION_CODE_ACCESS_TOKEN) {
      accessToken = context.secrets.OAUTH2_AUTHORIZATION_CODE_ACCESS_TOKEN;
    } else if (context.secrets?.OAUTH2_CLIENT_CREDENTIALS_CLIENT_SECRET) {
      const tokenUrl = context.environment?.OAUTH2_CLIENT_CREDENTIALS_TOKEN_URL;
      const clientId = context.environment?.OAUTH2_CLIENT_CREDENTIALS_CLIENT_ID;
      const clientSecret = context.secrets.OAUTH2_CLIENT_CREDENTIALS_CLIENT_SECRET;

      if (!tokenUrl || !clientId || !clientSecret) {
        throw new Error('OAuth2 Client Credentials flow requires TOKEN_URL, CLIENT_ID, and CLIENT_SECRET');
      }

      accessToken = await getClientCredentialsToken({
        tokenUrl,
        clientId,
        clientSecret,
        scope: context.environment?.OAUTH2_CLIENT_CREDENTIALS_SCOPE,
        audience: context.environment?.OAUTH2_CLIENT_CREDENTIALS_AUDIENCE,
        authStyle: context.environment?.OAUTH2_CLIENT_CREDENTIALS_AUTH_STYLE
      });
    } else {
      throw new Error('OAuth2 authentication is required. Configure either Authorization Code or Client Credentials flow.');
    }

    const { userPrincipalName, groupId } = params;

    console.log(`Removing user ${userPrincipalName} from group ${groupId}`);

    // Step 1: Get user's directory object ID
    console.log(`Step 1: Getting user directory object ID for ${userPrincipalName}`);
    const getUserResponse = await getUser(userPrincipalName, address, accessToken);

    if (!getUserResponse.ok) {
      throw new Error(`Failed to get user ${userPrincipalName}: ${getUserResponse.status} ${getUserResponse.statusText}`);
    }

    const userData = await getUserResponse.json();
    const userId = userData.id;

    if (!userId) {
      throw new Error(`No directory object ID found for user ${userPrincipalName}`);
    }

    console.log(`Found user directory object ID: ${userId}`);

    // Step 2: Remove user from group
    console.log(`Step 2: Removing user ${userId} from group ${groupId}`);
    const removeResponse = await removeUserFromGroup(groupId, userId, address, accessToken);

    // Handle success cases: 204 No Content or 404 Not Found (user not in group)
    if (removeResponse.status === 204) {
      console.log(`Successfully removed user ${userPrincipalName} from group ${groupId}`);
      return {
        status: 'success',
        userPrincipalName,
        groupId,
        userId,
        removed: true
      };
    } else if (removeResponse.status === 404) {
      console.log(`User ${userPrincipalName} was not a member of group ${groupId}`);
      return {
        status: 'success',
        userPrincipalName,
        groupId,
        userId,
        removed: false
      };
    } else {
      throw new Error(`Failed to remove user from group: ${removeResponse.status} ${removeResponse.statusText}`);
    }
  },

  /**
   * Error recovery handler - implements retry logic for transient failures
   * @param {Object} params - Original params plus error information
   * @param {Object} context - Execution context
   * @returns {Object} Recovery results
   */
  error: async (params, _context) => {
    const { error } = params;
    console.error(`Azure AD remove from group action encountered error: ${error.message}`);

    // Handle rate limiting - wait and let framework retry
    if (error.message.includes('429')) {
      console.log('Rate limited, waiting before retry');
      await new Promise(resolve => setTimeout(resolve, 5000));
      return { status: 'retry_requested' };
    }

    // Handle transient server errors - let framework retry
    if (error.message.includes('502') || error.message.includes('503') || error.message.includes('504')) {
      console.log('Server error encountered, requesting retry');
      return { status: 'retry_requested' };
    }

    // Authentication/authorization errors are fatal
    if (error.message.includes('401') || error.message.includes('403')) {
      console.error('Authentication/authorization error - not retryable');
      throw error;
    }

    // Default: let framework retry
    return { status: 'retry_requested' };
  },

  /**
   * Graceful shutdown handler - implements cleanup logic
   * @param {Object} params - Original params plus halt reason
   * @param {Object} context - Execution context
   * @returns {Object} Cleanup results
   */
  halt: async (params, _context) => {
    const { reason, userPrincipalName, groupId } = params;
    console.log(`Azure AD remove from group action halted (${reason}) for user ${userPrincipalName || 'unknown'} and group ${groupId || 'unknown'}`);

    return {
      status: 'halted',
      userPrincipalName: userPrincipalName || 'unknown',
      groupId: groupId || 'unknown',
      reason,
      halted_at: new Date().toISOString()
    };
  }
};