declare module "bwip-js" {
  export interface RenderOptions {
    bcid: string;
    text: string;
    scale?: number;
    width?: number;
    height?: number;
    includetext?: boolean;
    textsize?: number;
    textmargin?: number;
    padding?: number;
    background?: string;
    [key: string]: unknown;
  }

  export function toBuffer(opts: RenderOptions): Promise<Buffer>;

  const bwip: { toBuffer: typeof toBuffer; [key: string]: unknown };
  export default bwip;
}
