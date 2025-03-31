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
import { CHART_TYPE, SHAPE_TYPE } from "./constant";

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
    FontFace: z.string().optional().describe("Font face"),
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
    try {
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
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error adding text to slide "${slideId}": ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "add-table",
  "Add a table to the specified slide",
  {
    slideId: z.string().describe("ID of the slide"),
    data: z.array(
      z
        .array(
          z.object({
            text: z.string(),
            options: z
              .object({
                align: z
                  .enum(["left", "center", "right"])
                  .optional()
                  .describe("Table alignment"),
                bold: z
                  .boolean()
                  .optional()
                  .default(false)
                  .describe("Bold text"),
                border: z
                  .object({
                    type: z
                      .enum(["none", "solid", "dash"])
                      .describe("Border type"),
                    pt: z.number().min(0).max(10).describe("Border thickness"),
                    color: z.string().describe("Border color(hex code)"),
                  })
                  .optional()
                  .describe("Border options"),
                color: z.string().optional().describe("Text color(hex code)"),
                fill: z
                  .object({
                    color: z.string().describe("Fill color hex code"),
                    transparency: z
                      .number()
                      .min(0)
                      .max(100)
                      .optional()
                      .describe("Transparency of the fill color"),
                  })
                  .optional()
                  .describe("Fill color"),
                fontSize: z
                  .number()
                  .optional()
                  .default(18)
                  .describe("Font size of the text"),
                fontFace: z.string().optional().describe("Font face"),
              })
              .optional()
              .describe("Options for the cell"),
          })
        )
        .describe("Data for the table")
    ),
    x: z.number().min(0).max(10).describe("X position of the table"),
    y: z.number().min(0).max(5.5).describe("Y position of the table"),
    w: z.number().min(0).max(10).describe("Width of the table"),
    h: z.number().min(0).max(5.5).describe("Height of the table"),
    colW: z
      .array(z.number().min(0).max(10))
      .optional()
      .describe("Each column widths"),
    rowH: z
      .array(z.number().min(0).max(5.5))
      .optional()
      .describe("Each row heights"),
    align: z
      .enum(["left", "center", "right"])
      .optional()
      .describe("Table alignment"),
    bold: z.boolean().optional().default(false).describe("Bold text"),
    border: z
      .object({
        type: z.enum(["none", "solid", "dash"]).describe("Border type"),
        pt: z.number().min(0).max(10).describe("Border thickness"),
        color: z.string().describe("Border color(hex code)"),
      })
      .optional()
      .describe("Border options"),
    color: z.string().optional().describe("Text color(hex code)"),
    fill: z
      .object({
        color: z.string().describe("Fill color hex code"),
        transparency: z
          .number()
          .min(0)
          .max(100)
          .optional()
          .describe("Transparency of the fill color"),
      })
      .optional()
      .describe("Fill color"),
    fontSize: z
      .number()
      .optional()
      .default(18)
      .describe("Font size of the text"),
    fontFace: z.string().optional().describe("Font face"),
  },
  async ({
    slideId,
    data,
    x,
    y,
    w,
    h,
    colW,
    rowH,
    align,
    bold,
    border,
    color,
    fill,
    fontSize,
    fontFace,
  }) => {
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
    try {
      slide.addTable(data, {
        x,
        y,
        w,
        h,
        colW,
        rowH,
        align,
        bold,
        border,
        color,
        fill,
        fontSize,
        fontFace,
      });

      return {
        content: [
          {
            type: "text",
            text: `Table added to slide "${slideId}".`,
          },
        ],
        isError: false,
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error adding table to slide "${slideId}": ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "add-shape",
  "Add a shape to the specified slide",
  {
    slideId: z.string().describe("ID of the slide"),
    shape: z.enum(SHAPE_TYPE as [string, ...string[]]).describe("Shape type"),
    x: z.number().min(0).max(10).describe("X position of the shape"),
    y: z.number().min(0).max(5.5).describe("Y position of the shape"),
    w: z.number().min(0).max(10).describe("Width of the shape"),
    h: z.number().min(0).max(5.5).describe("Height of the shape"),
    align: z
      .enum(["left", "center", "right"])
      .optional()
      .describe("Shape alignment"),
    flipH: z.boolean().optional().default(false).describe("Flip horizontally"),
    flipV: z.boolean().optional().default(false).describe("Flip vertically"),
    line: z
      .object({
        color: z.string().describe("Line color(hex code)"),
        dashType: z
          .enum([
            "dash",
            "dashDot",
            "lgDash",
            "lgDashDot",
            "lgDashDotDot",
            "solid",
            "sysDash",
            "sysDot",
          ])
          .describe("Line type"),
        beginArrowType: z
          .enum(["arrow", "diamond", "oval", "stealth", "triangle", "none"])
          .describe("Line ending"),
        endArrowType: z
          .enum(["arrow", "diamond", "oval", "stealth", "triangle", "none"])
          .describe("Line heading"),
        transparency: z.number().min(0).max(100).describe("Line transparency"),
        width: z.number().min(1).max(256).describe("Line width"),
      })
      .optional()
      .describe("Shape Line options"),
    rectRadius: z
      .number()
      .min(0)
      .max(1)
      .optional()
      .describe("Rectangle radius(Only for roundRect)"),
    rotate: z.number().min(-360).max(360).optional().describe("Rotation angle"),
    fill: z
      .object({
        color: z.string().describe("Fill color hex code"),
        transparency: z
          .number()
          .min(0)
          .max(100)
          .optional()
          .describe("Transparency of the fill color"),
      })
      .optional()
      .describe("Fill color options"),
  },
  async ({ slideId, shape, x, y, w, h, fill }) => {
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
    try {
      slide.addShape(shape as pptxgen.SHAPE_NAME, {
        x,
        y,
        w,
        h,
        fill,
      });

      return {
        content: [
          {
            type: "text",
            text: `Shape added to slide "${slideId}".`,
          },
        ],
        isError: false,
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error adding shape to slide "${slideId}": ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "add-chart",
  "Add a chart to the specified slide",
  {
    slideId: z.string().describe("ID of the slide"),
    chartType: z
      .enum(CHART_TYPE as [string, ...string[]])
      .describe("Chart type"),
    data: z
      .array(
        z.object({
          name: z.string().describe("Chart name"),
          labels: z.array(z.string()).describe("Chart labels"),
          values: z.array(z.number()).describe("Chart values"),
        })
      )
      .describe("Chart data"),
    x: z.number().min(0).max(10).describe("X position of the chart"),
    y: z.number().min(0).max(5.5).describe("Y position of the chart"),
    w: z.number().min(0).max(10).describe("Width of the chart"),
    h: z.number().min(0).max(5.5).describe("Height of the chart"),
  },
  async ({ slideId, chartType, data, x, y, w, h }) => {
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
    try {
      slide.addChart(chartType as pptxgen.CHART_NAME, data, {
        x,
        y,
        w,
        h,
      });
      return {
        content: [
          {
            type: "text",
            text: `Chart added to slide "${slideId}".`,
          },
        ],
        isError: false,
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error adding chart to slide "${slideId}": ${error}`,
          },
        ],
        isError: true,
      };
    }
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
