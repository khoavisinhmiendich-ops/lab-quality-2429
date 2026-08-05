declare module 'mammoth' {
  export interface ConvertResult {
    value: string;
    messages: Array<{
      type: string;
      message: string;
    }>;
  }

  export function convertToHtml(options: {
    arrayBuffer: ArrayBuffer;
  }): Promise<ConvertResult>;
}