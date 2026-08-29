import test from 'node:test';
import assert from 'node:assert/strict';
import { buildReportDocument } from '../lib/report-generator.js';

test('buildReportDocument creates a valid document structure', () => {
  const doc = buildReportDocument({
    studentName: 'Jane Doe',
    degreeProgram: 'BS Information Systems',
    acknowledgement: ['Thank you for your support.'],
    organizationBackground: 'Sample organization background.',
    vision: 'To be the best.',
    mission: 'To serve the community.',
    objectives: ['Improve systems.'],
    coreValues: ['Integrity.'],
    services: ['IT support.'],
    strengths: ['Strong team.'],
    weaknesses: ['Limited resources.'],
    opportunities: ['Modernization.'],
    threats: ['Cybersecurity risks.'],
    recommendations: ['Upgrade systems.'],
    tasks: ['Maintain systems.'],
    procedures: ['Follow ticketing policy.'],
    issue1Description: 'Network downtime occurred.',
    issue1Strategy: 'Re-routed network traffic.',
    issue2Description: 'Legacy database failed.',
    issue2Strategy: 'Fixed SQL syntax and drivers.',
    lessons: ['Be systematic.'],
    selfEvaluation: ['I learned a lot.'],
    relevancy: ['This aligns with my studies.']
  });

  assert.ok(doc);
  assert.equal(typeof doc.add, 'function');
  assert.equal(typeof doc.Body, 'object');
});
