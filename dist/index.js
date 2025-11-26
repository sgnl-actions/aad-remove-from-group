// SGNL Job Script - Auto-generated bundle
'use strict';

/**
 * Azure Active Directory Remove User from Group Action
 *
 * This action removes a user from a group in Azure Active Directory using a two-step process:
 * 1. Get the user's directory object ID using their userPrincipalName
 * 2. Remove the user from the group using the directory object ID
 */

/**
 * Helper function to get a user by userPrincipalName
 * @param {string} userPrincipalName - User Principal Name (email)
 * @param {string} address - Azure AD base URL
 * @param {string} authToken - Bearer authentication token
 * @returns {Promise<Response>} HTTP response
 */
async function getUser(userPrincipalName, address, authToken) {
  // Remove trailing slash from address if present
  const cleanAddress = address.endsWith('/') ? address.slice(0, -1) : address;

  const encodedUPN = encodeURIComponent(userPrincipalName);
  const url = `${cleanAddress}/users/${encodedUPN}`;

  const response = await fetch(url, {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${authToken}`,
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
 * @param {string} authToken - Bearer authentication token
 * @returns {Promise<Response>} HTTP response
 */
async function removeUserFromGroup(groupId, userId, address, authToken) {
  // Remove trailing slash from address if present
  const cleanAddress = address.endsWith('/') ? address.slice(0, -1) : address;

  const encodedUserId = encodeURIComponent(userId);
  const url = `${cleanAddress}/groups/${groupId}/members/${encodedUserId}/$ref`;

  const response = await fetch(url, {
    method: 'DELETE',
    headers: {
      'Authorization': `Bearer ${authToken}`,
      'Accept': 'application/json'
    }
  });

  return response;
}

var script = {
  /**
   * Main execution handler - removes a user from an Azure AD group
   * @param {Object} params - Job input parameters
   * @param {string} params.userPrincipalName - User Principal Name (email) to remove from group
   * @param {string} params.groupId - Azure AD Group ID to remove user from
   * @param {string} params.address - The Azure AD API base URL (e.g., https://graph.microsoft.com)
   * @param {Object} context - Execution context with env, secrets, outputs
   * @param {string} context.environment.ADDRESS - Default Azure AD API base URL
   * @param {string} context.secrets.BEARER_AUTH_TOKEN - Bearer token for Azure AD API authentication
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

    if (!context.secrets.BEARER_AUTH_TOKEN) {
      throw new Error('BEARER_AUTH_TOKEN secret is required');
    }

    // Determine the URL to use
    const address = params.address || context.environment?.ADDRESS;
    if (!address) {
      throw new Error('No URL specified. Provide either address parameter or ADDRESS environment variable');
    }

    const { userPrincipalName, groupId } = params;
    const authToken = context.secrets.BEARER_AUTH_TOKEN;

    console.log(`Removing user ${userPrincipalName} from group ${groupId}`);

    // Step 1: Get user's directory object ID
    console.log(`Step 1: Getting user directory object ID for ${userPrincipalName}`);
    const getUserResponse = await getUser(userPrincipalName, address, authToken);

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
    const removeResponse = await removeUserFromGroup(groupId, userId, address, authToken);

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

module.exports = script;
