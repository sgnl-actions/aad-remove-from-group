/**
 * Azure Active Directory Remove User from Group Action
 *
 * This action removes a user from a group in Azure Active Directory using a two-step process:
 * 1. Get the user's directory object ID using their userPrincipalName
 * 2. Remove the user from the group using the directory object ID
 */

import { getBaseURL, createAuthHeaders } from '@sgnl-actions/utils';

/**
 * Helper function to get a user by userPrincipalName
 * @param {string} userPrincipalName - User Principal Name (email)
 * @param {string} baseUrl - Azure AD base URL
 * @param {Object} headers - Request headers with Authorization
 * @returns {Promise<Response>} HTTP response
 */
async function getUser(userPrincipalName, baseUrl, headers) {
  const encodedUPN = encodeURIComponent(userPrincipalName);
  const url = `${baseUrl}/v1.0/users/${encodedUPN}`;

  const response = await fetch(url, {
    method: 'GET',
    headers
  });

  return response;
}

/**
 * Helper function to check if user is already a member of the group
 * @param {string} userPrincipalName - User Principal Name (UPN) of the user
 * @param {string} groupId - Azure AD Group ID (GUID)
 * @param {string} baseUrl - Azure AD base URL
 * @param {Object} headers - Request headers with Authorization
 * @returns {Promise<boolean>} - True if user is a member, false otherwise
 */
async function isUserInGroup(userPrincipalName, groupId, baseUrl, headers) {
  const encodedUPN = encodeURIComponent(userPrincipalName);

  // Construct URL with proper encoding of the $filter parameter
  const baseURL = `${baseUrl}/v1.0/users/${encodedUPN}/memberOf`;
  const url = new URL(baseURL);
  url.searchParams.set('$filter', `id eq '${groupId}'`);

  const response = await fetch(url.toString(), {
    method: 'GET',
    headers
  });

  if (!response.ok) {
    throw new Error(`Failed to check group membership: ${response.status} ${response.statusText}`);
  }

  const data = await response.json();
  return data.value && data.value.length > 0;
}

/**
 * Helper function to remove a user from a group
 * @param {string} groupId - Azure AD Group ID
 * @param {string} userId - User's directory object ID
 * @param {string} baseUrl - Azure AD base URL
 * @param {Object} headers - Request headers with Authorization
 * @returns {Promise<Response>} HTTP response
 */
async function removeUserFromGroup(groupId, userId, baseUrl, headers) {
  const encodedUserId = encodeURIComponent(userId);
  const url = `${baseUrl}/v1.0/groups/${groupId}/members/${encodedUserId}/$ref`;

  const response = await fetch(url, {
    method: 'DELETE',
    headers
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
   *
   * @returns {Object} Job results
   */
  invoke: async (params, context) => {
    console.log('Starting Azure AD remove user from group action');

    const { userPrincipalName, groupId } = params;

    if (!userPrincipalName || typeof userPrincipalName !== 'string' || !userPrincipalName.trim()) {
      throw new Error('userPrincipalName parameter is required and cannot be empty');
    }

    if (!groupId || typeof groupId !== 'string' || !groupId.trim()) {
      throw new Error('groupId parameter is required and cannot be empty');
    }

    // Get base URL and authentication headers using utilities
    const baseUrl = getBaseURL(params, context);
    const headers = await createAuthHeaders(context);

    console.log(`Processing removal of user ${userPrincipalName} from group ${groupId}`);

    try {
      // Step 1: Check if user is actually a member of the group
      console.log(`Step 1: Checking if user ${userPrincipalName} is in group ${groupId}`);
      const isMember = await isUserInGroup(
        userPrincipalName,
        groupId,
        baseUrl,
        headers
      );

      if (!isMember) {
        console.log(`User ${userPrincipalName} is not a member of group ${groupId}`);
        return {
          status: 'success',
          userPrincipalName,
          groupId,
          removed: false,
          message: 'User is not a member of the group',
          address: baseUrl
        };
      }

      // Step 2: Get user's directory object ID (needed for removal API)
      console.log(`Step 2: Getting user directory object ID for ${userPrincipalName}`);
      const getUserResponse = await getUser(userPrincipalName, baseUrl, headers);

      if (!getUserResponse.ok) {
        throw new Error(`Failed to get user ${userPrincipalName}: ${getUserResponse.status} ${getUserResponse.statusText}`);
      }

      const userData = await getUserResponse.json();
      const userId = userData.id;

      if (!userId) {
        throw new Error(`No directory object ID found for user ${userPrincipalName}`);
      }

      console.log(`Found user directory object ID: ${userId}`);

      // Step 3: Remove user from group
      console.log(`Step 3: Removing user ${userId} from group ${groupId}`);
      const removeResponse = await removeUserFromGroup(groupId, userId, baseUrl, headers);

      if (removeResponse.status === 204) {
        console.log(`Successfully removed user ${userPrincipalName} from group ${groupId}`);
        return {
          status: 'success',
          userPrincipalName,
          groupId,
          userId,
          removed: true,
          address: baseUrl
        };
      } else {
        const errorText = await removeResponse.text();
        throw new Error(`Failed to remove user from group: ${removeResponse.status} ${removeResponse.statusText} - ${errorText}`);
      }
    } catch (error) {
      console.error(`Error in group membership removal operation: ${error.message}`);
      throw error;
    }
  },

  /**
   * Error recovery handler - framework handles retries by default
   * Only implement if custom recovery logic is needed
   * @param {Object} params - Original params plus error information
   * @param {Object} context - Execution context
   * @returns {Object} Recovery results
   */
  error: async (params, _context) => {
    const { error, userPrincipalName, groupId } = params;
    console.error(`User group removal failed for user ${userPrincipalName} from group ${groupId}: ${error.message}`);

    // Framework handles retries for transient errors (429, 502, 503, 504)
    // Just re-throw the error to let the framework handle it
    throw error;
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