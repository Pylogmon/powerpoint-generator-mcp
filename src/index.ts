#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import portfinder from "portfinder";
import { mkdirSync, existsSync } from "fs";
import express from "express";
import { nanoid } from "nanoid";
import pptxgen from "pptxgenjs";
import path from "path";
import { z } from "zod";
import os from "os";

// Global instance of pptxgen
let instances: Record<string, pptxgen> = {};
let slides: Record<string, pptxgen.Slide> = {};
let file_server_port = 60000;
const BASE_PORT = 60000;
const MAX_PORT = 65535;
const FILES_DIR = path.join(os.tmpdir(), "mcp-powerpoint-generator");

portfinder.basePort = BASE_PORT;
portfinder.highestPort = MAX_PORT;

portfinder.getPort((err, port) => {
  if (err) {
    console.error("Could not find an open port:", err);
    process.exit(1);
  }
  file_server_port = port;
  const app = express();

  if (!existsSync(FILES_DIR)) {
    mkdirSync(FILES_DIR, { recursive: true });
  }
  app.use("/", express.static(FILES_DIR));

  // 启动服务器
  app.listen(port, () => {
    console.log(`File Server running on port ${port}`);
  });
});

// Create server instance
const server = new McpServer({
  name: "mcp-powerpoint-generator",
  version: "0.1.0",
  capabilities: {
    resources: {},
    tools: {},
  },
});

// Define the tool to create a PowerPoint presentation
server.tool(
  "create-presentation",
  "Create a new PowerPoint presentation (Allways call this tool first and only once)",
  {
    title: z
      .string()
      .optional()
      .default("PowerPoint Presentation")
      .describe("Title of the presentation"),
    subject: z
      .string()
      .optional()
      .default("PowerPoint MCP Server")
      .describe("Subject of the presentation"),
    author: z
      .string()
      .optional()
      .default("PowerPoint MCP Server")
      .describe("Author of the presentation"),
    company: z
      .string()
      .optional()
      .default("PowerPoint MCP Server")
      .describe("Company of the presentation"),
    revision: z
      .string()
      .optional()
      .default("1")
      .describe("Revision of the presentation"),
    layout: z
      .enum(["LAYOUT_16x9", "LAYOUT_16x10", "LAYOUT_4x3", "LAYOUT_WIDE"])
      .optional()
      .default("LAYOUT_16x9")
      .describe("Layout of the presentation"),
    rtl: z.boolean().default(false).describe("Right to left layout"),
  },
  async ({ title, subject, author, company, revision, layout, rtl }) => {
    const id = nanoid();
    let pptx = new pptxgen();
    pptx.title = title;
    pptx.subject = subject;
    pptx.author = author;
    pptx.company = company;
    pptx.revision = revision;
    pptx.layout = layout;
    pptx.rtlMode = rtl;

    instances[id] = pptx;

    return {
      content: [
        {
          type: "text",
          text: `PowerPoint presentation "${id}" created.`,
        },
      ],
      isError: false,
    };
  }
);

// Define the tool to add a slide to the presentation
server.tool(
  "add-slide",
  "Add a slide to the PowerPoint presentation",
  {
    id: z.string().describe("ID of the presentation"),
  },
  async ({ id }) => {
    if (!instances[id]) {
      return {
        content: [
          {
            type: "text",
            text: `PowerPoint presentation "${id}" not found, please create it first.`,
          },
        ],
        isError: true,
      };
    }
    const slideId = nanoid();
    let pptx = instances[id];
    let slide = pptx.addSlide();
    slides[slideId] = slide;

    return {
      content: [
        {
          type: "text",
          text: `Slide "${slideId}" added to presentation "${id}".`,
        },
      ],
      isError: false,
    };
  }
);
// Define the tool to add a text box to the slide
server.tool(
  "add-text",
  "Add a text box to the slide",
  {
    slideId: z.string().describe("ID of the slide"),
    text: z.string().describe("Text to add"),
    x: z.number().default(0).describe("X position of the text box"),
    y: z.number().default(0).describe("Y position of the text box"),
    w: z.number().default(8).describe("Width of the text box"),
    h: z.number().default(1).describe("Height of the text box"),
  },
  async ({ slideId, text, x, y, w, h }) => {
    if (!slides[slideId]) {
      return {
        content: [
          {
            type: "text",
            text: `Slide "${slideId}" not found, please create it first.`,
          },
        ],
        isError: true,
      };
    }
    let slide = slides[slideId];
    slide.addText(text, { x, y, w, h });

    return {
      content: [
        {
          type: "text",
          text: `Text box added to slide "${slideId}".`,
        },
      ],
      isError: false,
    };
  }
);

// Define the tool to save the presentation
server.tool(
  "save-presentation",
  "Save the PowerPoint presentation (Always call this tool last), After this tool is executed, you must send the returned URL to the user.",
  {
    id: z
      .string()
      .describe(
        "The ID of the presentation returned by the create-presentation tool"
      ),
  },
  async ({ id }) => {
    if (!instances[id]) {
      return {
        content: [
          {
            type: "text",
            text: `PowerPoint presentation "${id}" not found, please create it first.`,
          },
        ],
        isError: true,
      };
    }
    let pptx = instances[id];
    const title = pptx.title || "PowerPoint Presentation";
    try {
      await pptx.writeFile({
        fileName: path.join(FILES_DIR, `${title}-${id}.pptx`),
      });
      delete instances[id];
      delete slides[id];
      return {
        content: [
          {
            type: "text",
            text: `PowerPoint presentation has saved and can be find at http://localhost:${file_server_port}/${encodeURIComponent(
              title
            )}-${id}.pptx`,
          },
        ],
        isError: false,
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error saving PowerPoint presentation "${id}": ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("PowerPoint MCP Server running on stdio");
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
});
