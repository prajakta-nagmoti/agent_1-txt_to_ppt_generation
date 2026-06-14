# Sample: Product Architecture Overview

This document describes the product architecture and rollout plan.

## Goals

- Launch MVP in Q3
- Validate core assumptions

## Architecture

```mermaid
flowchart TD
    A[Client] --> B[API Gateway]
    B --> C{Services}
    C --> D[Auth Service]
    C --> E[Data Service]
```

## Roadmap

- Phase 1: Core features
- Phase 2: Scaling and monitoring

## Summary

This deck summarizes goals, architecture and roadmap.