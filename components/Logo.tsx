'use client';

import React, { useId } from 'react';

interface LeafSpec {
  x: number;
  y: number;
  angle: number;
  scale: number;
}

// Hand-tuned points tracing the wreath's arc from the crossed base up each side —
// mirrored at render time so only one side needs to be authored.
const LEFT_LEAVES: LeafSpec[] = [
  { x: 74, y: 182, angle: 255, scale: 1.0 },
  { x: 58, y: 170, angle: 235, scale: 1.05 },
  { x: 45, y: 152, angle: 210, scale: 1.1 },
  { x: 37, y: 128, angle: 185, scale: 1.12 },
  { x: 36, y: 102, angle: 160, scale: 1.1 },
  { x: 43, y: 78, angle: 135, scale: 1.02 },
  { x: 57, y: 59, angle: 110, scale: 0.92 },
  { x: 76, y: 47, angle: 88, scale: 0.8 },
];

const LEAF_D = 'M0,-9.5 C3.4,-6 3.4,4.5 0,9.5 C-3.4,4.5 -3.4,-6 0,-9.5 Z';

const Leaf: React.FC<{ leaf: LeafSpec; index: number; mirrored?: boolean }> = ({ leaf, index, mirrored }) => (
  // Positioning (translate/rotate/scale-to-size) lives on the static wrapper group so the
  // CSS entrance animation on the inner path — which only touches opacity/scale — never
  // has to fight over the `transform` property and overwrite the leaf's placement.
  <g transform={`translate(${leaf.x} ${leaf.y}) rotate(${mirrored ? 180 - leaf.angle : leaf.angle}) scale(${leaf.scale})`}>
    <path
      d={LEAF_D}
      fill="url(#hlogo-leafGradient)"
      className="hlogo-leaf"
      style={{ animationDelay: `${80 + index * 55}ms` }}
    />
  </g>
);

export const HospitalLogo: React.FC<{ className?: string; animated?: boolean }> = ({
  className = 'w-10 h-10',
  animated = true,
}) => {
  const rawId = useId();
const uid = rawId.replace(/:/g, ''); // Loại bỏ ký tự ':' nếu SVG ID không thích ứng

  return (
    <svg
      viewBox="0 0 200 200"
      className={`${className} ${animated ? 'hlogo-root' : ''}`}
      xmlns="http://www.w3.org/2000/svg"
      role="img"
      aria-label="Huy hiệu Bệnh viện Phong - Da liễu TW Quy Hòa"
    >
      <style>{`
        .hlogo-root .hlogo-glow { animation: hlogo-glow 3.2s ease-in-out infinite; }
        .hlogo-root .hlogo-ring { animation: hlogo-ringIn 0.7s cubic-bezier(0.16,1,0.3,1) both; }
        .hlogo-root .hlogo-leaf { animation: hlogo-leafIn 0.55s cubic-bezier(0.34,1.56,0.64,1) both; transform-origin: center; }
        .hlogo-root .hlogo-emblem { animation: hlogo-riseIn 0.6s cubic-bezier(0.16,1,0.3,1) both; animation-delay: 0.55s; }
        .hlogo-root .hlogo-snake { animation: hlogo-draw 0.9s ease-out both; animation-delay: 0.65s; }
        .hlogo-root .hlogo-text { animation: hlogo-riseIn 0.5s ease-out both; animation-delay: 1.05s; }
        .hlogo-root:hover .hlogo-mark { transform: scale(1.035); }
        .hlogo-mark { transform-origin: 100px 100px; transition: transform 0.35s cubic-bezier(0.16,1,0.3,1); }

        @keyframes hlogo-glow {
          0%, 100% { opacity: 0.35; r: 92; }
          50% { opacity: 0.6; r: 96; }
        }
        @keyframes hlogo-ringIn {
          from { opacity: 0; stroke-dashoffset: 600; }
          to { opacity: 1; stroke-dashoffset: 0; }
        }
        @keyframes hlogo-leafIn {
          from { opacity: 0; transform: scale(0); }
          to { opacity: 1; transform: scale(1); }
        }
        @keyframes hlogo-riseIn {
          from { opacity: 0; transform: translateY(6px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes hlogo-draw {
          from { stroke-dashoffset: 220; }
          to { stroke-dashoffset: 0; }
        }
        @media (prefers-reduced-motion: reduce) {
          .hlogo-root .hlogo-glow,
          .hlogo-root .hlogo-ring,
          .hlogo-root .hlogo-leaf,
          .hlogo-root .hlogo-emblem,
          .hlogo-root .hlogo-snake,
          .hlogo-root .hlogo-text { animation: none !important; opacity: 1 !important; }
          .hlogo-mark { transition: none !important; }
        }
      `}</style>

      <defs>
        <linearGradient id={`hlogo-red-${uid}`} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor="#F0453D" />
          <stop offset="100%" stopColor="#D6281F" />
        </linearGradient>
        <linearGradient id="hlogo-leafGradient" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor="#4CC26B" />
          <stop offset="100%" stopColor="#22A34A" />
        </linearGradient>
      </defs>

      {/* Soft ambient glow ring */}
      <circle className="hlogo-glow" cx="100" cy="100" r="92" fill="none" stroke="#E53E3E" strokeWidth="10" opacity="0.35" />

      <g className="hlogo-mark">
        {/* Outer red ring */}
        <circle
          className="hlogo-ring"
          cx="100"
          cy="100"
          r="92"
          fill="#FFFFFF"
          stroke={`url(#hlogo-red-${uid})`}
          strokeWidth="4.5"
          strokeDasharray="600"
        />

        {/* Laurel wreath, left authored + right mirrored */}
        <g>
          {LEFT_LEAVES.map((leaf, i) => (
            <Leaf key={`l-${i}`} leaf={leaf} index={i} />
          ))}
        </g>
        <g transform="translate(200,0) scale(-1,1)">
          {LEFT_LEAVES.map((leaf, i) => (
            <Leaf key={`r-${i}`} leaf={leaf} index={i} mirrored />
          ))}
        </g>
        {/* Crossed stems at the base of the wreath */}
        <path
          d="M76,183 Q100,192 124,183"
          fill="none"
          stroke="#22A34A"
          strokeWidth="4"
          strokeLinecap="round"
          className="hlogo-leaf"
          style={{ animationDelay: '520ms' }}
        />

        {/* Central emblem: staff, wings, snake, base, text */}
        <g className="hlogo-emblem">
          {/* Winged crossbar */}
          <path d="M62,63 L138,63 L109,93 L91,93 Z" fill={`url(#hlogo-red-${uid})`} />

          {/* Staff shaped like a stylised P */}
          <path
            d="M92,63 L92,151 L108,151 L108,96 C132,96 132,63 108,63 Z"
            fill={`url(#hlogo-red-${uid})`}
          />

          {/* Snake coiling the staff */}
          <path
            d="M83,150 C112,138 68,122 100,108 C132,94 74,79 100,64"
            fill="none"
            stroke={`url(#hlogo-red-${uid})`}
            strokeWidth="5.5"
            strokeLinecap="round"
            strokeDasharray="220"
            className="hlogo-snake"
          />
          <path
            d="M100,64 L94,55 M100,64 L104,54"
            stroke={`url(#hlogo-red-${uid})`}
            strokeWidth="2.5"
            strokeLinecap="round"
          />

          {/* Base plinth */}
          <rect x="76" y="151" width="48" height="10" rx="1.5" fill={`url(#hlogo-red-${uid})`} />
          <rect x="66" y="177" width="68" height="3" rx="1.5" fill={`url(#hlogo-red-${uid})`} />
        </g>

        {/* Institution monogram */}
        <text
          x="100"
          y="172"
          textAnchor="middle"
          fill={`url(#hlogo-red-${uid})`}
          fontSize="14.5"
          fontWeight={800}
          fontFamily="'Inter', Arial, sans-serif"
          letterSpacing="0.5"
          className="hlogo-text"
        >
          QUYHOANDH
        </text>
      </g>
    </svg>
  );
};

export default HospitalLogo;
