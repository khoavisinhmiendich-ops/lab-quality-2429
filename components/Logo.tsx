'use client';

import React from 'react';

export const HospitalLogo: React.FC<{ className?: string }> = ({ className = "w-10 h-10" }) => {
  return (
    <svg viewBox="0 0 200 200" className={className} xmlns="http://www.w3.org/2000/svg">
      <circle cx="100" cy="100" r="95" fill="none" stroke="#E53E3E" strokeWidth="4" />
      <g fill="none" stroke="#22C55E" strokeWidth="4" strokeLinecap="round">
        <path d="M 40,140 Q 25,100 45,65 Q 60,95 40,140 Z" />
        <path d="M 50,155 Q 35,120 60,85 Q 70,115 50,155 Z" />
        <path d="M 65,168 Q 50,135 75,105" />
        <path d="M 160,140 Q 175,100 155,65 Q 140,95 160,140 Z" />
        <path d="M 150,155 Q 165,120 140,85 Q 130,115 150,155 Z" />
        <path d="M 135,168 Q 150,135 125,105" />
        <path d="M 70,175 Q 100,190 130,175" />
      </g>
      <path d="M 60,65 L 140,65 L 110,95 L 90,95 Z" fill="#E53E3E" />
      <rect x="92" y="95" width="16" height="55" fill="#E53E3E" />
      <rect x="75" y="142" width="50" height="10" fill="#E53E3E" />
      <path d="M 98,135 Q 85,120 98,105 Q 112,90 98,75" fill="none" stroke="#E53E3E" strokeWidth="3" />
      <text x="100" y="125" textAnchor="middle" fill="#E53E3E" fontSize="13" fontWeight="bold" fontFamily="sans-serif">
        QUYHOANDH
      </text>
    </svg>
  );
};