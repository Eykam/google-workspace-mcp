/**
 * @license
 * Copyright 2025 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import { google, slides_v1, drive_v3 } from 'googleapis';
import * as fs from 'node:fs/promises';
import { createReadStream } from 'node:fs';
import * as path from 'node:path';
import { request } from 'gaxios';
import { AuthManager } from '../auth/AuthManager';
import { logToFile } from '../utils/logger';
import { extractDocId } from '../utils/IdUtils';
import { gaxiosOptions } from '../utils/GaxiosConfig';
import { buildDriveSearchQuery, MIME_TYPES } from '../utils/DriveQueryBuilder';

export class SlidesService {
  constructor(private authManager: AuthManager) {}

  private async getSlidesClient(): Promise<slides_v1.Slides> {
    const auth = await this.authManager.getAuthenticatedClient();
    const options = { ...gaxiosOptions, auth };
    return google.slides({ version: 'v1', ...options });
  }

  private async getDriveClient(): Promise<drive_v3.Drive> {
    const auth = await this.authManager.getAuthenticatedClient();
    const options = { ...gaxiosOptions, auth };
    return google.drive({ version: 'v3', ...options });
  }

  public getText = async ({ presentationId }: { presentationId: string }) => {
    logToFile(
      `[SlidesService] Starting getText for presentation: ${presentationId}`,
    );
    try {
      const id = extractDocId(presentationId) || presentationId;

      const slides = await this.getSlidesClient();
      // Get the presentation with all necessary fields
      const presentation = await slides.presentations.get({
        presentationId: id,
        fields:
          'title,slides(pageElements(shape(text,shapeProperties),table(tableRows(tableCells(text)))))',
      });

      let content = '';

      // Add presentation title
      if (presentation.data.title) {
        content += `Presentation Title: ${presentation.data.title}\n\n`;
      }

      // Process each slide
      if (presentation.data.slides) {
        presentation.data.slides.forEach((slide, slideIndex) => {
          content += `\n--- Slide ${slideIndex + 1} ---\n`;

          if (slide.pageElements) {
            slide.pageElements.forEach((element) => {
              // Extract text from shapes
              if (element.shape && element.shape.text) {
                const shapeText = this.extractTextFromTextContent(
                  element.shape.text,
                );
                if (shapeText) {
                  content += shapeText + '\n';
                }
              }

              // Extract text from tables
              if (element.table && element.table.tableRows) {
                content += '\n--- Table Data ---\n';
                element.table.tableRows.forEach((row) => {
                  const rowText: string[] = [];
                  if (row.tableCells) {
                    row.tableCells.forEach((cell) => {
                      const cellText = cell.text
                        ? this.extractTextFromTextContent(cell.text)
                        : '';
                      rowText.push(cellText.trim());
                    });
                  }
                  content += rowText.join(' | ') + '\n';
                });
                content += '--- End Table Data ---\n';
              }
            });
          }
          content += '\n';
        });
      }

      logToFile(`[SlidesService] Finished getText for presentation: ${id}`);
      return {
        content: [
          {
            type: 'text' as const,
            text: content.trim(),
          },
        ],
      };
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : String(error);
      logToFile(`[SlidesService] Error during slides.getText: ${errorMessage}`);
      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({ error: errorMessage }),
          },
        ],
      };
    }
  };

  private extractTextFromTextContent(
    textContent: slides_v1.Schema$TextContent,
  ): string {
    let text = '';
    if (textContent.textElements) {
      textContent.textElements.forEach((element) => {
        if (element.textRun && element.textRun.content) {
          text += element.textRun.content;
        } else if (element.paragraphMarker) {
          // Add newline for paragraph markers
          text += '\n';
        }
      });
    }
    return text;
  }

  public find = async ({
    query,
    pageToken,
    pageSize = 10,
  }: {
    query: string;
    pageToken?: string;
    pageSize?: number;
  }) => {
    logToFile(
      `[SlidesService] Searching for presentations with query: ${query}`,
    );
    try {
      const q = buildDriveSearchQuery(MIME_TYPES.PRESENTATION, query);
      logToFile(`[SlidesService] Executing Drive API query: ${q}`);

      const drive = await this.getDriveClient();
      const res = await drive.files.list({
        pageSize: pageSize,
        fields: 'nextPageToken, files(id, name)',
        q: q,
        pageToken: pageToken,
      });

      const files = res.data.files || [];
      const nextPageToken = res.data.nextPageToken;

      logToFile(`[SlidesService] Found ${files.length} presentations.`);

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({
              files: files,
              nextPageToken: nextPageToken,
            }),
          },
        ],
      };
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : String(error);
      logToFile(`[SlidesService] Error during slides.find: ${errorMessage}`);
      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({ error: errorMessage }),
          },
        ],
      };
    }
  };

  public getMetadata = async ({
    presentationId,
  }: {
    presentationId: string;
  }) => {
    logToFile(
      `[SlidesService] Starting getMetadata for presentation: ${presentationId}`,
    );
    try {
      const id = extractDocId(presentationId) || presentationId;

      const slides = await this.getSlidesClient();
      const presentation = await slides.presentations.get({
        presentationId: id,
        fields:
          'presentationId,title,slides(objectId,pageElements(objectId,table(rows,columns))),pageSize,notesMaster,masters,layouts',
      });

      const metadata = {
        presentationId: presentation.data.presentationId,
        title: presentation.data.title,
        slideCount: presentation.data.slides?.length || 0,
        // The tool description points callers here for slide object ids, and
        // slides.getSlideThumbnail is unusable without them.
        slideObjectIds: (presentation.data.slides || []).map((s) => s.objectId),
        // Table element ids, so callers can target tables with batchUpdate.
        tables: (presentation.data.slides || []).flatMap((s) =>
          (s.pageElements || [])
            .filter((el) => !!el.table)
            .map((el) => ({
              slideObjectId: s.objectId,
              objectId: el.objectId,
              rows: el.table?.rows,
              columns: el.table?.columns,
            })),
        ),
        pageSize: presentation.data.pageSize,
        hasMasters: !!presentation.data.masters?.length,
        hasLayouts: !!presentation.data.layouts?.length,
        hasNotesMaster: !!presentation.data.notesMaster,
      };

      logToFile(`[SlidesService] Finished getMetadata for presentation: ${id}`);
      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(metadata),
          },
        ],
      };
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : String(error);
      logToFile(
        `[SlidesService] Error during slides.getMetadata: ${errorMessage}`,
      );
      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({ error: errorMessage }),
          },
        ],
      };
    }
  };

  private async downloadToLocal(url: string, localPath: string) {
    logToFile(`[SlidesService] Downloading from ${url} to ${localPath}`);
    if (!path.isAbsolute(localPath)) {
      throw new Error('localPath must be an absolute path.');
    }

    // Ensure directory exists
    await fs.mkdir(path.dirname(localPath), { recursive: true });

    const response = await request({
      url,
      responseType: 'arraybuffer',
      ...gaxiosOptions,
    });

    await fs.writeFile(localPath, Buffer.from(response.data as ArrayBuffer));
    logToFile(`[SlidesService] Downloaded successfully to ${localPath}`);
    return localPath;
  }

  public getImages = async ({
    presentationId,
    localPath,
  }: {
    presentationId: string;
    localPath: string;
  }) => {
    logToFile(
      `[SlidesService] Starting getImages for presentation: ${presentationId} (localPath: ${localPath})`,
    );
    try {
      const id = extractDocId(presentationId) || presentationId;
      const slides = await this.getSlidesClient();
      const presentation = await slides.presentations.get({
        presentationId: id,
        fields:
          'slides(objectId,pageElements(objectId,title,description,image(contentUrl,sourceUrl)))',
      });

      const images = await Promise.all(
        (presentation.data.slides ?? []).flatMap((slide, index) =>
          (slide.pageElements ?? [])
            .filter((element) => element.image)
            .map(async (element) => {
              const imageData: any = {
                slideIndex: index + 1,
                slideObjectId: slide.objectId,
                elementObjectId: element.objectId,
                title: element.title,
                description: element.description,
                contentUrl: element.image?.contentUrl,
                sourceUrl: element.image?.sourceUrl,
              };

              if (imageData.contentUrl) {
                const filename = `slide_${imageData.slideIndex}_${element.objectId}.png`;
                const fullPath = path.join(localPath, filename);
                try {
                  await this.downloadToLocal(imageData.contentUrl, fullPath);
                  imageData.localPath = fullPath;
                } catch (downloadError) {
                  logToFile(
                    `[SlidesService] Failed to download image ${element.objectId}: ${downloadError}`,
                  );
                  imageData.downloadError = String(downloadError);
                }
              }

              return imageData;
            }),
        ),
      );

      logToFile(`[SlidesService] Finished getImages for presentation: ${id}`);
      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({ images }),
          },
        ],
      };
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : String(error);
      logToFile(
        `[SlidesService] Error during slides.getImages: ${errorMessage}`,
      );
      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({ error: errorMessage }),
          },
        ],
      };
    }
  };

  public getSlideThumbnail = async ({
    presentationId,
    slideObjectId,
    localPath,
  }: {
    presentationId: string;
    slideObjectId: string;
    localPath: string;
  }) => {
    logToFile(
      `[SlidesService] Starting getSlideThumbnail for presentation: ${presentationId}, slide: ${slideObjectId} (localPath: ${localPath})`,
    );
    try {
      const id = extractDocId(presentationId) || presentationId;
      const slides = await this.getSlidesClient();
      const thumbnail = await slides.presentations.pages.getThumbnail({
        presentationId: id,
        pageObjectId: slideObjectId,
      });

      const result: any = { ...thumbnail.data };

      if (result.contentUrl) {
        try {
          await this.downloadToLocal(result.contentUrl, localPath);
          result.localPath = localPath;
        } catch (downloadError) {
          logToFile(
            `[SlidesService] Failed to download thumbnail for slide ${slideObjectId}: ${downloadError}`,
          );
          result.downloadError = String(downloadError);
        }
      }

      logToFile(
        `[SlidesService] Finished getSlideThumbnail for slide: ${slideObjectId}`,
      );
      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result),
          },
        ],
      };
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : String(error);
      logToFile(
        `[SlidesService] Error during slides.getSlideThumbnail: ${errorMessage}`,
      );
      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({ error: errorMessage }),
          },
        ],
      };
    }
  };

  private ok(payload: unknown) {
    return {
      content: [{ type: 'text' as const, text: JSON.stringify(payload) }],
    };
  }

  private fail(where: string, error: unknown) {
    const message = error instanceof Error ? error.message : String(error);
    logToFile(`[SlidesService] Error during ${where}: ${message}`);
    return {
      content: [
        { type: 'text' as const, text: JSON.stringify({ error: message }) },
      ],
    };
  }

  /**
   * Writes go through the Slides API, which accepts the full Drive scope this
   * server already requests, so no consent-screen change is needed.
   */
  public create = async ({ title }: { title: string }) => {
    logToFile(`[SlidesService] Creating presentation: ${title}`);
    try {
      const slides = await this.getSlidesClient();
      const res = await slides.presentations.create({
        requestBody: { title },
      });
      const id = res.data.presentationId;
      return this.ok({
        presentationId: id,
        title: res.data.title,
        url: `https://docs.google.com/presentation/d/${id}/edit`,
        slides: (res.data.slides || []).map((s) => s.objectId),
      });
    } catch (error) {
      return this.fail('slides.create', error);
    }
  };

  public batchUpdate = async ({
    presentationId,
    requests,
  }: {
    presentationId: string;
    requests: string | unknown[];
  }) => {
    try {
      const id = extractDocId(presentationId) || presentationId;
      const parsed: slides_v1.Schema$Request[] =
        typeof requests === 'string'
          ? JSON.parse(requests)
          : (requests as slides_v1.Schema$Request[]);
      if (!Array.isArray(parsed)) {
        throw new Error('requests must be a JSON array of Slides API requests');
      }
      logToFile(
        `[SlidesService] batchUpdate on ${id} with ${parsed.length} requests`,
      );
      const slides = await this.getSlidesClient();
      const res = await slides.presentations.batchUpdate({
        presentationId: id,
        requestBody: { requests: parsed },
      });
      return this.ok({
        presentationId: id,
        applied: parsed.length,
        replies: res.data.replies,
      });
    } catch (error) {
      return this.fail('slides.batchUpdate', error);
    }
  };

  /**
   * Uploads a local .pptx and lets Drive convert it to a native Google Slides
   * presentation, which is the only reliable way to land a deck built offline.
   */
  public importPptx = async ({
    localPath,
    title,
    parentId,
  }: {
    localPath: string;
    title?: string;
    parentId?: string;
  }) => {
    logToFile(`[SlidesService] Importing pptx from ${localPath}`);
    try {
      const stats = await fs.stat(localPath);
      if (!stats.isFile()) {
        throw new Error(`Not a file: ${localPath}`);
      }
      const drive = await this.getDriveClient();
      const res = await drive.files.create({
        requestBody: {
          name: title || path.basename(localPath, path.extname(localPath)),
          mimeType: 'application/vnd.google-apps.presentation',
          ...(parentId ? { parents: [parentId] } : {}),
        },
        media: {
          mimeType:
            'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          body: createReadStream(localPath),
        },
        supportsAllDrives: true,
        fields: 'id,name,webViewLink',
      });
      return this.ok({
        presentationId: res.data.id,
        title: res.data.name,
        url:
          res.data.webViewLink ||
          `https://docs.google.com/presentation/d/${res.data.id}/edit`,
        bytesUploaded: stats.size,
      });
    } catch (error) {
      return this.fail('slides.importPptx', error);
    }
  };

  /**
   * Replaces the content of an existing presentation from a local .pptx while
   * keeping the same file id, so links already shared stay valid.
   */
  public updateFromPptx = async ({
    presentationId,
    localPath,
  }: {
    presentationId: string;
    localPath: string;
  }) => {
    logToFile(`[SlidesService] Replacing ${presentationId} from ${localPath}`);
    try {
      const id = extractDocId(presentationId) || presentationId;
      const stats = await fs.stat(localPath);
      if (!stats.isFile()) {
        throw new Error(`Not a file: ${localPath}`);
      }
      const drive = await this.getDriveClient();
      const res = await drive.files.update({
        fileId: id,
        media: {
          mimeType:
            'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          body: createReadStream(localPath),
        },
        supportsAllDrives: true,
        fields: 'id,name,webViewLink,modifiedTime',
      });
      return this.ok({
        presentationId: res.data.id,
        title: res.data.name,
        url:
          res.data.webViewLink ||
          `https://docs.google.com/presentation/d/${res.data.id}/edit`,
        modifiedTime: res.data.modifiedTime,
        bytesUploaded: stats.size,
      });
    } catch (error) {
      return this.fail('slides.updateFromPptx', error);
    }
  };
}
