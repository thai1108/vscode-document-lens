import { compoundFileHeaderStub } from '../support/compoundFile.ts';
import type { Sample } from '../support/samples.ts';

export const samples: Sample[] = [{ fileName: 'sample.doc', build: () => compoundFileHeaderStub() }];
