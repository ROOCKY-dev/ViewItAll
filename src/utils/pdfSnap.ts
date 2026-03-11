import type { SnapDirection } from '../settings';

export type Point = { x: number; y: number };

/**
 * Constrain `raw` to the given snap direction relative to `origin`.
 * All coordinates are normalised 0-1 fractions of the canvas.
 */
export function snapPoint(origin: Point, raw: Point, direction: SnapDirection): Point {
	const dx = raw.x - origin.x;
	const dy = raw.y - origin.y;

	switch (direction) {
		case 'horizontal':
			return { x: raw.x, y: origin.y };

		case 'vertical':
			return { x: origin.x, y: raw.y };

		case 'slope': {
			// Snap to nearest 45° increment (8 directions)
			const angle   = Math.atan2(dy, dx);
			const snapped = Math.round(angle / (Math.PI / 4)) * (Math.PI / 4);
			const dist    = Math.sqrt(dx * dx + dy * dy);
			return {
				x: origin.x + dist * Math.cos(snapped),
				y: origin.y + dist * Math.sin(snapped),
			};
		}
	}
}
