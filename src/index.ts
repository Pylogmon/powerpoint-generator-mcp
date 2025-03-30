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

  app.listen(port, () => {
    console.log(`File Server running on port ${port}`);
  });
});

const server = new McpServer({
  name: "mcp-powerpoint-generator",
  version: "0.1.1",
  capabilities: {
    resources: {},
    tools: {},
  },
});

server.tool(
  "create-presentation",
  "Create a new PowerPoint presentation",
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
    rtl: z.boolean().default(false).describe("Right to left layout"),
  },
  async ({ title, subject, author, company, revision, rtl }) => {
    const id = nanoid();
    let pptx = new pptxgen();
    pptx.title = title;
    pptx.subject = subject;
    pptx.author = author;
    pptx.company = company;
    pptx.revision = revision;
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

server.tool(
  "add-text",
  "Add text to the specified slide",
  {
    slideId: z.string().describe("ID of the slide"),
    text: z.string().describe("Text to add"),
    x: z.number().min(0).max(10).describe("X position of the text box"),
    y: z.number().min(0).max(5.5).describe("Y position of the text box"),
    w: z.number().min(0).max(10).describe("Width of the text box"),
    h: z.number().min(0).max(5.5).describe("Height of the text box"),
    align: z
      .enum(["left", "center", "right", "justify"])
      .optional()
      .describe("Text alignment"),
    bold: z.boolean().optional().default(false).describe("Bold text"),
    color: z.string().optional().describe("Text color(hex code)"),
    fontSize: z
      .number()
      .optional()
      .default(18)
      .describe("Font size of the text"),
  },
  async ({ slideId, text, x, y, w, h, align, bold, color, fontSize }) => {
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
    slide.addText(text, { x, y, w, h, align, bold, color, fontSize });

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

server.tool(
  "get-file-url",
  "Get the presentation URL, use this tool when you need to get a download link of presentation after completing the presentation.",
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
            text: `http://localhost:${file_server_port}/${encodeURIComponent(
              title
            )}-${id}.pptx`,
          },
          {
            type: "text",
            text: `NOTE: The user cannot see this link, you need to explicitly send this link to the user.It is recommended to send using the following format: [Download ${title}.pptx](http://localhost:${file_server_port}/${encodeURIComponent(
              title
            )}-${id}.pptx)`,
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
