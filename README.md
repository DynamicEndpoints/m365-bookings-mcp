# Microsoft 365 Bookings MCP Server
[![smithery badge](https://smithery.ai/badge/@DynamicEndpoints/m365-bookings-mcp)](https://smithery.ai/server/@DynamicEndpoints/m365-bookings-mcp)

An MCP server that provides tools for interacting with Microsoft Bookings through the Microsoft Graph API.

## Features

- List Bookings businesses
- Get staff members for a business
- Get services offered by a business
- Get appointments for a business

## Setup

### Installing via Smithery

To install Microsoft 365 Bookings for Claude Desktop automatically via [Smithery](https://smithery.ai/server/@DynamicEndpoints/m365-bookings-mcp):

```bash
npx -y @smithery/cli install @DynamicEndpoints/m365-bookings-mcp --client claude
```

### Manual Installation
1. Create an Azure AD application registration:
   - Go to Azure Portal > Azure Active Directory > App registrations
   - Create a new registration
   - Add Microsoft Graph API permissions:
     - BookingsAppointment.ReadWrite.All
     - BookingsBusiness.ReadWrite.All
     - BookingsStaffMember.ReadWrite.All

2. Create a .env file with the following variables:
```
MICROSOFT_GRAPH_CLIENT_ID=your-client-id
MICROSOFT_GRAPH_CLIENT_SECRET=your-client-secret
MICROSOFT_GRAPH_TENANT_ID=your-tenant-id
```

3. Install dependencies:
```bash
npm install
```

4. Build the server:
```bash
npm run build
```

## Available Tools

### get_bookings_businesses
Get a list of all Bookings businesses in the organization.

### get_business_staff
Get staff members for a specific Bookings business.
- Required parameter: businessId

### get_business_services
Get services offered by a specific Bookings business.
- Required parameter: businessId

### get_business_appointments
Get appointments for a specific Bookings business.
- Required parameter: businessId
- Optional parameters: 
  - startDate (ISO format)
  - endDate (ISO format)
