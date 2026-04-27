import { useEffect, useMemo, useState } from 'react';
import type { IListViewModeConfig } from '../config/types';
import { UsersService } from '../../../../services/users/UsersService';
import {
  collectDistinctAccessWebKeys,
  normWebPath,
} from './viewModeAccess';

export interface IViewModeMembershipState {
  userId: number;
  groupByWeb: Map<string, Set<number>>;
  pageNorm: string;
}

export function useViewModeMembership(
  viewModes: IListViewModeConfig[],
  pageWebServerRelativeUrl: string | undefined
): IViewModeMembershipState | null {
  const pageNorm = useMemo(() => normWebPath(pageWebServerRelativeUrl || '/'), [pageWebServerRelativeUrl]);
  const [state, setState] = useState<IViewModeMembershipState | null>(null);

  useEffect(() => {
    const us = new UsersService();
    let cancelled = false;
    const needsMembership = viewModes.some((m) => m.access !== undefined && m.access !== null);
    const needsGroups = viewModes.some((m) => (m.access?.allowedGroupIds?.length ?? 0) > 0);

    us.getCurrentUser()
      .then(async (u) => {
        if (cancelled) return;
        if (!needsMembership) {
          setState({ userId: u.Id, groupByWeb: new Map(), pageNorm });
          return;
        }
        if (!needsGroups) {
          setState({ userId: u.Id, groupByWeb: new Map(), pageNorm });
          return;
        }
        const keys = collectDistinctAccessWebKeys(viewModes, pageNorm);
        const pairs = await Promise.all(
          keys.map(async (k) => {
            const ids = await us.getCurrentUserGroupIds(k === pageNorm ? undefined : k);
            return [k, new Set(ids)] as const;
          })
        );
        if (cancelled) return;
        setState({ userId: u.Id, groupByWeb: new Map(pairs), pageNorm });
      })
      .catch(() => {
        if (!cancelled) setState({ userId: 0, groupByWeb: new Map(), pageNorm });
      });
    return () => {
      cancelled = true;
    };
  }, [viewModes, pageNorm]);

  return state;
}
