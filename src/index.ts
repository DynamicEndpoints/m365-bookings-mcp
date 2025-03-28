#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';

const {
  MICROSOFT_GRAPH_CLIENT_ID,
  MICROSOFT_GRAPH_CLIENT_SECRET,
  MICROSOFT_GRAPH_TENANT_ID,
} = process.env;

if (!MICROSOFT_GRAPH_CLIENT_ID || !MICROSOFT_GRAPH_CLIENT_SECRET || !MICROSOFT_GRAPH_TENANT_ID) {
  throw new Error('Missing required environment variables');
}

class BookingsServer {
  private server: Server;
  private graphClient!: Client; // Using definite assignment assertion

  constructor() {
    this.server = new Server(
      {
        name: 'm365-bookings',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    // Initialize Graph client
    this.initializeGraphClient();
    
    // Set up request handlers
    this.setupHandlers();
    
    // Error handling
    this.server.onerror = (error) => console.error('[MCP Error]', error);
    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  private async initializeGraphClient() {
    // Get access token using client credentials flow
    const tokenResponse = await fetch(`https://login.microsoftonline.com/${MICROSOFT_GRAPH_TENANT_ID}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: new URLSearchParams({
        client_id: MICROSOFT_GRAPH_CLIENT_ID || '',
        client_secret: MICROSOFT_GRAPH_CLIENT_SECRET || '',
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials',
      }),
    });

    const { access_token } = await tokenResponse.json();

    this.graphClient = Client.init({
      authProvider: (done) => {
        done(null, access_token);
      },
    });
  }

  private setupHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: 'get_bookings_businesses',
          description: 'Get list of Bookings businesses',
          inputSchema: {
            type: 'object',
            properties: {},
          },
        },
        {
          name: 'get_business_staff',
          description: 'Get staff members for a Bookings business',
          inputSchema: {
            type: 'object',
            properties: {
              businessId: {
                type: 'string',
                description: 'ID of the Bookings business',
              },
            },
            required: ['businessId'],
          },
        },
        {
          name: 'get_business_services',
          description: 'Get services offered by a Bookings business',
          inputSchema: {
            type: 'object',
            properties: {
              businessId: {
                type: 'string',
                description: 'ID of the Bookings business',
              },
            },
            required: ['businessId'],
          },
        },
        {
          name: 'get_business_appointments',
          description: 'Get appointments for a Bookings business',
          inputSchema: {
            type: 'object',
            properties: {
              businessId: {
                type: 'string',
                description: 'ID of the Bookings business',
              },
              startDate: {
                type: 'string',
                description: 'Start date for appointments (ISO format)',
              },
              endDate: {
                type: 'string',
                description: 'End date for appointments (ISO format)',
              },
            },
            required: ['businessId'],
          },
        },
      ],
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      try {
        switch (request.params.name) {
          case 'get_bookings_businesses':
            return await this.getBookingsBusinesses();
          case 'get_business_staff': {
            const args = request.params.arguments as { businessId: string };
            return await this.getBusinessStaff(args.businessId);
          }
          case 'get_business_services': {
            const args = request.params.arguments as { businessId: string };
            return await this.getBusinessServices(args.businessId);
          }
          case 'get_business_appointments': {
            const args = request.params.arguments as {
              businessId: string;
              startDate?: string;
              endDate?: string;
            };
            return await this.getBusinessAppointments(
              args.businessId,
              args.startDate,
              args.endDate
            );
          }
          default:
            throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${request.params.name}`);
        }
      } catch (error: any) {
        console.error('Error executing tool:', error);
        return {
          content: [
            {
              type: 'text',
              text: `Error: ${error.message}`,
            },
          ],
          isError: true,
        };
      }
    });
  }

  private async getBookingsBusinesses() {
    const response = await this.graphClient
      .api('/solutions/bookingBusinesses')
      .get();

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(response.value, null, 2),
        },
      ],
    };
  }

  private async getBusinessStaff(businessId: string) {
    const response = await this.graphClient
      .api(`/solutions/bookingBusinesses/${businessId}/staffMembers`)
      .get();

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(response.value, null, 2),
        },
      ],
    };
  }

  private async getBusinessServices(businessId: string) {
    const response = await this.graphClient
      .api(`/solutions/bookingBusinesses/${businessId}/services`)
      .get();

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(response.value, null, 2),
        },
      ],
    };
  }

  private async getBusinessAppointments(businessId: string, startDate?: string, endDate?: string) {
    let api = this.graphClient
      .api(`/solutions/bookingBusinesses/${businessId}/appointments`);

    if (startDate || endDate) {
      const filter = [];
      if (startDate) filter.push(`start ge ${startDate}`);
      if (endDate) filter.push(`end le ${endDate}`);
      api = api.filter(filter.join(' and '));
    }

    const response = await api.get();

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(response.value, null, 2),
        },
      ],
    };
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Microsoft 365 Bookings MCP server running on stdio');
  }
}

const server = new BookingsServer();
server.run().catch(console.error);
