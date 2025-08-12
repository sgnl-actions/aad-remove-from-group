import script from '../src/script.mjs';

// Mock fetch globally
const mockFetch = jest.fn();
global.fetch = mockFetch;

describe('Azure AD Remove from Group Script', () => {
  const mockContext = {
    environment: {
      AZURE_AD_TENANT_URL: 'https://graph.microsoft.com/v1.0/'
    },
    secrets: {
      AZURE_AD_TOKEN: 'test-token-123456'
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
            'Accept': 'application/json'
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
            'Accept': 'application/json'
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

    test('should throw error if userPrincipalName is missing', async () => {
      const params = {
        groupId: 'group-123-456-789'
      };

      await expect(script.invoke(params, mockContext)).rejects.toThrow('userPrincipalName is required');
    });

    test('should throw error if groupId is missing', async () => {
      const params = {
        userPrincipalName: 'test@example.com'
      };

      await expect(script.invoke(params, mockContext)).rejects.toThrow('groupId is required');
    });

    test('should throw error if AZURE_AD_TOKEN secret is missing', async () => {
      const params = {
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789'
      };

      const contextWithoutToken = {
        ...mockContext,
        secrets: {}
      };

      await expect(script.invoke(params, contextWithoutToken)).rejects.toThrow('AZURE_AD_TOKEN secret is required');
    });

    test('should throw error if AZURE_AD_TENANT_URL environment is missing', async () => {
      const params = {
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789'
      };

      const contextWithoutTenantUrl = {
        ...mockContext,
        environment: {}
      };

      await expect(script.invoke(params, contextWithoutTenantUrl)).rejects.toThrow('AZURE_AD_TENANT_URL environment variable is required');
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
    test('should handle rate limiting (429) with retry request', async () => {
      const params = {
        error: { message: 'Rate limited: 429' },
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789'
      };

      // Mock setTimeout to resolve immediately for testing
      jest.spyOn(global, 'setTimeout').mockImplementation((fn) => fn());

      const result = await script.error(params, mockContext);

      expect(result.status).toBe('retry_requested');
      expect(setTimeout).toHaveBeenCalledWith(expect.any(Function), 5000);

      setTimeout.mockRestore();
    });

    test('should handle server errors (502, 503, 504) with retry request', async () => {
      const serverErrors = ['502', '503', '504'];

      for (const errorCode of serverErrors) {
        const params = {
          error: { message: `Server error: ${errorCode}` },
          userPrincipalName: 'test@example.com',
          groupId: 'group-123-456-789'
        };

        const result = await script.error(params, mockContext);
        expect(result.status).toBe('retry_requested');
      }
    });

    test('should throw error for authentication failures (401, 403)', async () => {
      // Test 401 error
      const params401 = {
        error: { message: 'Auth error: 401' },
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789'
      };

      let thrown401 = false;
      try {
        await script.error(params401, mockContext);
      } catch (error) {
        thrown401 = true;
        expect(error.message).toContain('Auth error: 401');
      }
      expect(thrown401).toBe(true);

      // Test 403 error
      const params403 = {
        error: { message: 'Auth error: 403' },
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789'
      };

      let thrown403 = false;
      try {
        await script.error(params403, mockContext);
      } catch (error) {
        thrown403 = true;
        expect(error.message).toContain('Auth error: 403');
      }
      expect(thrown403).toBe(true);
    });

    test('should request retry for other errors', async () => {
      const params = {
        error: { message: 'Some other error' },
        userPrincipalName: 'test@example.com',
        groupId: 'group-123-456-789'
      };

      const result = await script.error(params, mockContext);
      expect(result.status).toBe('retry_requested');
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
});