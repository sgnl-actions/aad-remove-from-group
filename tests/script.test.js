import script from '../src/script.mjs';
import { SGNL_USER_AGENT } from '@sgnl-actions/utils';

// Mock fetch globally
const mockFetch = jest.fn();
global.fetch = mockFetch;

describe('Azure AD Remove from Group Script', () => {
  const mockContext = {
    environment: {
      ADDRESS: 'https://graph.microsoft.com'
    },
    secrets: {
      OAUTH2_AUTHORIZATION_CODE_ACCESS_TOKEN: 'test-token-123456'
    }
  };

  const mockUserData = {
    id: 'user-123-456-789',
    userPrincipalName: 'test@example.com'
  };

  beforeEach(() => {
    jest.clearAllMocks();
    console.log = jest.fn();
    console.error = jest.fn();
  });

  describe('invoke handler', () => {
    test('should successfully remove user from group', async () => {
      const params = {
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789'
      };

      // Mock successful user lookup
      mockFetch
        .mockResolvedValueOnce({
          ok: true,
          json: async () => mockUserData
        })
        // Mock successful group removal (204 No Content)
        .mockResolvedValueOnce({
          ok: true,
          status: 204
        });

      const result = await script.invoke(params, mockContext);

      expect(result.status).toBe('success');
      expect(result.userPrincipalName).toBe('test@example.com');
      expect(result.groupId).toBe('group-123-456-789');
      expect(result.userId).toBe('user-123-456-789');
      expect(result.removed).toBe(true);

      // Verify both API calls were made
      expect(mockFetch).toHaveBeenCalledTimes(2);

      // Check user lookup call
      expect(mockFetch).toHaveBeenNthCalledWith(1,
        'https://graph.microsoft.com/v1.0/users/test%40example.com',
        {
          method: 'GET',
          headers: {
            'Authorization': 'Bearer test-token-123456',
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'User-Agent': SGNL_USER_AGENT
          }
        }
      );

      // Check group removal call
      expect(mockFetch).toHaveBeenNthCalledWith(2,
        'https://graph.microsoft.com/v1.0/groups/group-123-456-789/members/user-123-456-789/$ref',
        {
          method: 'DELETE',
          headers: {
            'Authorization': 'Bearer test-token-123456',
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'User-Agent': SGNL_USER_AGENT
          }
        }
      );
    });

    test('should handle user not in group (404 response)', async () => {
      const params = {
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789'
      };

      // Mock successful user lookup
      mockFetch
        .mockResolvedValueOnce({
          ok: true,
          json: async () => mockUserData
        })
        // Mock 404 Not Found (user not in group)
        .mockResolvedValueOnce({
          ok: false,
          status: 404
        });

      const result = await script.invoke(params, mockContext);

      expect(result.status).toBe('success');
      expect(result.userPrincipalName).toBe('test@example.com');
      expect(result.groupId).toBe('group-123-456-789');
      expect(result.userId).toBe('user-123-456-789');
      expect(result.removed).toBe(false);
    });

    test('should handle URL encoding for userPrincipalName with special characters', async () => {
      const params = {
        userPrincipalName: 'test+user@example.com',
        groupId: 'group-123-456-789'
      };

      // Mock successful user lookup
      mockFetch
        .mockResolvedValueOnce({
          ok: true,
          json: async () => ({ ...mockUserData, userPrincipalName: 'test+user@example.com' })
        })
        // Mock successful group removal
        .mockResolvedValueOnce({
          ok: true,
          status: 204
        });

      await script.invoke(params, mockContext);

      // Check that userPrincipalName was URL encoded
      expect(mockFetch).toHaveBeenNthCalledWith(1,
        'https://graph.microsoft.com/v1.0/users/test%2Buser%40example.com',
        expect.any(Object)
      );
    });

    test('should handle URL encoding for userId in group removal', async () => {
      const params = {
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789'
      };

      const userWithSpecialChars = {
        ...mockUserData,
        id: 'user+123&456=789'
      };

      // Mock successful user lookup
      mockFetch
        .mockResolvedValueOnce({
          ok: true,
          json: async () => userWithSpecialChars
        })
        // Mock successful group removal
        .mockResolvedValueOnce({
          ok: true,
          status: 204
        });

      await script.invoke(params, mockContext);

      // Check that userId was URL encoded in group removal call
      expect(mockFetch).toHaveBeenNthCalledWith(2,
        'https://graph.microsoft.com/v1.0/groups/group-123-456-789/members/user%2B123%26456%3D789/$ref',
        expect.any(Object)
      );
    });





    test('should throw error if user lookup fails', async () => {
      const params = {
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789'
      };

      // Mock failed user lookup
      mockFetch.mockResolvedValueOnce({
        ok: false,
        status: 404,
        statusText: 'Not Found'
      });

      await expect(script.invoke(params, mockContext)).rejects.toThrow('Failed to get user test@example.com: 404 Not Found');
    });

    test('should throw error if user has no directory object ID', async () => {
      const params = {
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789'
      };

      // Mock user lookup with missing ID
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({ userPrincipalName: 'test@example.com' }) // No id field
      });

      await expect(script.invoke(params, mockContext)).rejects.toThrow('No directory object ID found for user test@example.com');
    });

    test('should throw error if group removal fails with unexpected status', async () => {
      const params = {
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789'
      };

      // Mock successful user lookup
      mockFetch
        .mockResolvedValueOnce({
          ok: true,
          json: async () => mockUserData
        })
        // Mock failed group removal
        .mockResolvedValueOnce({
          ok: false,
          status: 403,
          statusText: 'Forbidden'
        });

      await expect(script.invoke(params, mockContext)).rejects.toThrow('Failed to remove user from group: 403 Forbidden');
    });
  });

  describe('error handler', () => {
    test('should re-throw error and let framework handle retries', async () => {
      const errorObj = new Error('Rate limited: 429');
      const params = {
        error: errorObj,
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789'
      };

      await expect(script.error(params, mockContext)).rejects.toThrow(errorObj);
      expect(console.error).toHaveBeenCalledWith(
        'User group removal failed for user test@example.com from group group-123-456-789: Rate limited: 429'
      );
    });

    test('should re-throw server errors', async () => {
      const errorObj = new Error('Server error: 502');
      const params = {
        error: errorObj,
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789'
      };

      await expect(script.error(params, mockContext)).rejects.toThrow(errorObj);
    });

    test('should re-throw authentication errors', async () => {
      const errorObj = new Error('Auth error: 401');
      const params = {
        error: errorObj,
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789'
      };

      await expect(script.error(params, mockContext)).rejects.toThrow(errorObj);
    });

    test('should re-throw any error', async () => {
      const errorObj = new Error('Some other error');
      const params = {
        error: errorObj,
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789'
      };

      await expect(script.error(params, mockContext)).rejects.toThrow(errorObj);
    });
  });

  describe('halt handler', () => {
    test('should handle graceful shutdown with all parameters', async () => {
      const params = {
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789',
        reason: 'timeout'
      };

      const result = await script.halt(params, mockContext);

      expect(result.status).toBe('halted');
      expect(result.userPrincipalName).toBe('test@example.com');
      expect(result.groupId).toBe('group-123-456-789');
      expect(result.reason).toBe('timeout');
      expect(result.halted_at).toBeDefined();
    });

    test('should handle halt without userPrincipalName or groupId', async () => {
      const params = {
        reason: 'system_shutdown'
      };

      const result = await script.halt(params, mockContext);

      expect(result.status).toBe('halted');
      expect(result.userPrincipalName).toBe('unknown');
      expect(result.groupId).toBe('unknown');
      expect(result.reason).toBe('system_shutdown');
    });
  });

  describe('invoke handler - idempotency', () => {
    test('should succeed on first call with removed:true', async () => {
      mockFetch
        .mockResolvedValueOnce({ ok: true, json: async () => mockUserData })
        .mockResolvedValueOnce({ ok: false, status: 204 });

      const result = await script.invoke({
        userPrincipalName: 'test@example.com',
        groupId: 'group-123'
      }, mockContext);

      expect(result.status).toBe('success');
      expect(result.removed).toBe(true);
    });

    test('should succeed on second call with removed:false (404 - user not in group)', async () => {
      mockFetch
        .mockResolvedValueOnce({ ok: true, json: async () => mockUserData })
        .mockResolvedValueOnce({ ok: false, status: 404 });

      const result = await script.invoke({
        userPrincipalName: 'test@example.com',
        groupId: 'group-123'
      }, mockContext);

      expect(result.status).toBe('success');
      expect(result.removed).toBe(false);
    });

    // Documents real Azure behavior: returns 400 with "modified properties: 'members'"
    // when removing a user who is not in the group
    test('should succeed on second call with removed:false (real Azure 400 response)', async () => {
      mockFetch
        .mockResolvedValueOnce({ ok: true, json: async () => mockUserData })
        .mockResolvedValueOnce({
          ok: false,
          status: 400,
          statusText: 'Bad Request',
          text: async () => JSON.stringify({
            error: {
              code: 'Request_BadRequest',
              message: "One or more removed object references do not exist for the following modified properties: 'members'."
            }
          })
        });

      const result = await script.invoke({
        userPrincipalName: 'test@example.com',
        groupId: 'group-123'
      }, mockContext);

      expect(result.status).toBe('success');
      expect(result.removed).toBe(false);
    });

    test('should produce same end state on repeated calls', async () => {
      mockFetch
        .mockResolvedValueOnce({ ok: true, json: async () => mockUserData })
        .mockResolvedValueOnce({ ok: true, status: 204 })
        .mockResolvedValueOnce({ ok: true, json: async () => mockUserData })
        .mockResolvedValueOnce({ ok: false, status: 404 });

      const r1 = await script.invoke({ userPrincipalName: 'test@example.com', groupId: 'group-123' }, mockContext);
      const r2 = await script.invoke({ userPrincipalName: 'test@example.com', groupId: 'group-123' }, mockContext);

      expect(r1.status).toBe('success');
      expect(r2.status).toBe('success');
      expect(r1.userPrincipalName).toBe(r2.userPrincipalName);
      expect(r1.groupId).toBe(r2.groupId);
      expect(r1.userId).toBe(r2.userId);
    });
  });

  describe('invoke handler - input validation', () => {
    test('should throw when userPrincipalName is missing', async () => {
      await expect(script.invoke({
        groupId: 'group-123'
      }, mockContext)).rejects.toThrow('userPrincipalName parameter is required and cannot be empty');

      expect(mockFetch).not.toHaveBeenCalled();
    });

    test('should throw when groupId is missing', async () => {
      await expect(script.invoke({
        userPrincipalName: 'test@example.com'
      }, mockContext)).rejects.toThrow('groupId parameter is required and cannot be empty');

      expect(mockFetch).not.toHaveBeenCalled();
    });

    test('should throw when auth token is missing', async () => {
      await expect(script.invoke({
        userPrincipalName: 'test@example.com',
        groupId: 'group-123'
      }, { environment: { ADDRESS: 'https://graph.microsoft.com' }, secrets: {} }))
        .rejects.toThrow(/No authentication configured/);

      expect(mockFetch).not.toHaveBeenCalled();
    });
  });

  describe('invoke handler - request construction', () => {
    test('should include User-Agent header in both API calls', async () => {
      mockFetch
        .mockResolvedValueOnce({ ok: true, json: async () => mockUserData })
        .mockResolvedValueOnce({ ok: true, status: 204 });

      await script.invoke({
        userPrincipalName: 'test@example.com',
        groupId: 'group-123'
      }, mockContext);

      for (const call of mockFetch.mock.calls) {
        expect(call[1].headers['User-Agent']).toBe(SGNL_USER_AGENT);
      }
    });

    test('should use custom address from params over environment ADDRESS', async () => {
      mockFetch
        .mockResolvedValueOnce({ ok: true, json: async () => mockUserData })
        .mockResolvedValueOnce({ ok: true, status: 204 });

      await script.invoke({
        userPrincipalName: 'test@example.com',
        groupId: 'group-123',
        address: 'https://custom-proxy.example.com'
      }, mockContext);

      expect(mockFetch).toHaveBeenNthCalledWith(1,
        expect.stringContaining('https://custom-proxy.example.com'),
        expect.any(Object)
      );
    });

    test('should use DELETE method for group removal', async () => {
      mockFetch
        .mockResolvedValueOnce({ ok: true, json: async () => mockUserData })
        .mockResolvedValueOnce({ ok: true, status: 204 });

      await script.invoke({
        userPrincipalName: 'test@example.com',
        groupId: 'group-123'
      }, mockContext);

      expect(mockFetch).toHaveBeenNthCalledWith(2,
        expect.any(String),
        expect.objectContaining({ method: 'DELETE' })
      );
    });

    test('should use GET method for user lookup', async () => {
      mockFetch
        .mockResolvedValueOnce({ ok: true, json: async () => mockUserData })
        .mockResolvedValueOnce({ ok: true, status: 204 });

      await script.invoke({
        userPrincipalName: 'test@example.com',
        groupId: 'group-123'
      }, mockContext);

      expect(mockFetch).toHaveBeenNthCalledWith(1,
        expect.any(String),
        expect.objectContaining({ method: 'GET' })
      );
    });
  });

  describe('invoke handler - network failures', () => {
    test('should throw when fetch rejects during user lookup', async () => {
      mockFetch.mockRejectedValueOnce(new Error('Network timeout'));

      await expect(script.invoke({
        userPrincipalName: 'test@example.com',
        groupId: 'group-123'
      }, mockContext)).rejects.toThrow('Network timeout');

      expect(mockFetch).toHaveBeenCalledTimes(1);
    });

    test('should throw when fetch rejects during group removal', async () => {
      mockFetch
        .mockResolvedValueOnce({ ok: true, json: async () => mockUserData })
        .mockRejectedValueOnce(new Error('Connection refused'));

      await expect(script.invoke({
        userPrincipalName: 'test@example.com',
        groupId: 'group-123'
      }, mockContext)).rejects.toThrow('Connection refused');

      expect(mockFetch).toHaveBeenCalledTimes(2);
    });
  });

  describe('invoke handler - error responses', () => {
    test('should throw on 401 during user lookup', async () => {
      mockFetch.mockResolvedValueOnce({
        ok: false, status: 401, statusText: 'Unauthorized'
      });

      await expect(script.invoke({
        userPrincipalName: 'test@example.com',
        groupId: 'group-123'
      }, mockContext)).rejects.toThrow(/401 Unauthorized/);
    });

    test('should throw on 403 during group removal', async () => {
      mockFetch
        .mockResolvedValueOnce({ ok: true, json: async () => mockUserData })
        .mockResolvedValueOnce({ ok: false, status: 403, statusText: 'Forbidden' });

      await expect(script.invoke({
        userPrincipalName: 'test@example.com',
        groupId: 'group-123'
      }, mockContext)).rejects.toThrow(/403 Forbidden/);
    });

    test('should throw on 500 during group removal', async () => {
      mockFetch
        .mockResolvedValueOnce({ ok: true, json: async () => mockUserData })
        .mockResolvedValueOnce({ ok: false, status: 500, statusText: 'Internal Server Error' });

      await expect(script.invoke({
        userPrincipalName: 'test@example.com',
        groupId: 'group-123'
      }, mockContext)).rejects.toThrow(/500 Internal Server Error/);
    });
  });

  describe('halt handler - edge cases', () => {
    test('should handle halt with no params at all', async () => {
      const result = await script.halt({}, mockContext);

      expect(result.status).toBe('halted');
      expect(result.userPrincipalName).toBe('unknown');
      expect(result.groupId).toBe('unknown');
      expect(result.reason).toBeUndefined();
      expect(result.halted_at).toBeDefined();
    });

    test('should include ISO timestamp in halted_at', async () => {
      const result = await script.halt({ reason: 'test' }, mockContext);
      expect(new Date(result.halted_at).toISOString()).toBe(result.halted_at);
    });
  });
});