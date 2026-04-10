import { graphfi, GraphFI, SPFx } from '@pnp/graph';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import '@pnp/graph/users';
import '@pnp/graph/groups';

let _graph: GraphFI;

export const getGraph = (context?: WebPartContext): GraphFI => {
  if (context) {
    _graph = graphfi().using(SPFx(context));
  }
  return _graph;
};


