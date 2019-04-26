import { WebPartContext } from '@microsoft/sp-webpart-base';
import { GraphDataOptions } from '../utils/GraphDataOptions';

export interface IGraphConnectionTestProps {
  selectedGraphDataOption: GraphDataOptions;
  context: WebPartContext;
}
